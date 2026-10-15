(function () {
  const DATA = window.__DATA__;
  const IS_ADMIN = !!window.__IS_ADMIN__;

  function getCookie(name) {
    let cookieValue = null;
    if (document.cookie && document.cookie !== "") {
      const cookies = document.cookie.split(";");
      for (let i = 0; i < cookies.length; i++) {
        const cookie = cookies[i].trim();
        if (cookie.substring(0, name.length + 1) === (name + "=")) {
          cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
          break;
        }
      }
    }
    return cookieValue;
  }
  const csrftoken = getCookie("csrftoken") || (document.querySelector("[name=csrfmiddlewaretoken]") ? document.querySelector("[name=csrfmiddlewaretoken]").value : "");

  const PRODUCTS = ["callable", "noncallable", "ib_callable", "ib_noncallable", "fin_callable"];
  const LABELS = {
    callable: "Non Fin Callable",
    noncallable: "Non Fin / Fin Non-Callable",
    ib_callable: "IB Callable",
    ib_noncallable: "IB Non-Callable",
    fin_callable: "Financial Callable",
  };

  let productKey = "callable";
  let saved = false;
  let isNominalMode = false;

  function toggleNominalMode() {
    isNominalMode = !isNominalMode;
    render();
  }

  function cur() {
    if (!DATA) return null;
    if (DATA[productKey]) return DATA[productKey];
    const keys = Object.keys(DATA);
    if (keys.length > 0) {
      productKey = keys[0];
      return DATA[productKey];
    }
    return null;
  }
  function deep(o) { return JSON.parse(JSON.stringify(o)); }
  function esc(s) {
    if (s == null) return "";
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function rkey(t, s) { return t + "|" + s; }
  function nid(prefix) { return prefix + Math.random().toString(36).slice(2, 7); }

  function parseAmt(v) {
    const s = String(v || "").replace(/[^\d.]/g, "");
    if (!s) return null;
    const n = parseFloat(s);
    return isNaN(n) ? null : n;
  }

  function fmtINR(n) {
    if (n == null || isNaN(n)) return "";
    return Number(n).toLocaleString("en-IN");
  }

  // ---------- validation ----------
  function validate(w) {
    const errors = [], badSlabs = new Set(), badTenures = new Set(), badRates = new Set();
    if (!w) return { errors, badSlabs, badTenures, badRates };
    const slabs = w.slabs || [];
    const tenures = w.tenures || [];
    const rates = w.rates || {};
    const iv = [];
    slabs.forEach((s) => {
      const f = parseAmt(s.amtFrom), t = parseAmt(s.amtTo), label = s.label || "(unnamed)";
      if (f == null || t == null) { errors.push('Slab "' + label + '" is missing an amount.'); badSlabs.add(s.id); }
      else if (f < 30000000) { errors.push('Slab "' + label + '": Amt From cannot be below 3 Cr (3,00,00,000).'); badSlabs.add(s.id); }
      else if (t > 99999999999) { errors.push('Slab "' + label + '": Amt To cannot exceed 99,99,99,99,999.'); badSlabs.add(s.id); }
      else if (f > t) { errors.push('Slab "' + label + '": Amt From is greater than Amt To.'); badSlabs.add(s.id); }
      else iv.push({ id: s.id, from: f, to: t, label });
    });
    iv.sort((a, b) => a.from - b.from);
    for (let i = 1; i < iv.length; i++) {
      if (iv[i].from <= iv[i - 1].to) {
        errors.push('Amounts overlap: "' + iv[i - 1].label + '" and "' + iv[i].label + '" cover the same range.');
        badSlabs.add(iv[i].id); badSlabs.add(iv[i - 1].id);
      }
    }
    for (let i = 0; i < tenures.length; i++) {
      const t = tenures[i];
      const tm = +t.fromMonth || 0, td = +t.toDays || 0;
      if (tm === 0 && td < 7) {
        errors.push("Tenure row " + (i + 1) + " (" + tm + "m " + td + "d): Duration cannot be below 7 days.");
        badTenures.add(t.id);
      } else if (tm > 120 || (tm === 120 && td > 0)) {
        errors.push("Tenure row " + (i + 1) + " (" + tm + "m " + td + "d): Duration cannot exceed 120 months.");
        badTenures.add(t.id);
      }
    }
    for (let i = 1; i < tenures.length; i++) {
      const a = tenures[i - 1], b = tenures[i];
      const am = +a.fromMonth || 0, ad = +a.toDays || 0, bm = +b.fromMonth || 0, bd = +b.toDays || 0;
      if (!(bm > am || (bm === am && bd > ad))) {
        errors.push("Tenure not ascending: " + bm + "m " + bd + "d must come after " + am + "m " + ad + "d.");
        badTenures.add(b.id);
        badTenures.add(a.id);
      }
    }
    tenures.forEach((t, tIdx) => {
      slabs.forEach((s, sIdx) => {
        const legacyKey = rkey(t.id, s.id);
        const idxKey = "idx_" + tIdx + "_" + sIdx;
        const val = rates[idxKey] !== undefined ? rates[idxKey] : (rates[legacyKey] || "");
        if (val === "" || val == null) {
          errors.push('Rate missing at Tenure row ' + (tIdx + 1) + ' (' + t.fromMonth + 'm ' + t.toDays + 'd), Slab "' + (s.label || sIdx + 1) + '".');
          badRates.add(legacyKey); badRates.add(idxKey);
        } else {
          const num = parseFloat(val);
          if (isNaN(num)) {
            errors.push('Rate "' + val + '" is not a valid number.');
            badRates.add(legacyKey); badRates.add(idxKey);
          } else if (num < 2 || num > 10) {
            errors.push('Rate "' + val + '%" must be between 2% and 10%.');
            badRates.add(legacyKey); badRates.add(idxKey);
          } else if (String(val).includes(".")) {
            const dec = String(val).split(".")[1] || "";
            if (dec.length > 2) {
              errors.push('Rate "' + val + '%" can have at most 2 decimal places.');
              badRates.add(legacyKey); badRates.add(idxKey);
            }
          }
        }
      });
    });
    return { errors, badSlabs, badTenures, badRates };
  }

  // ---------- grid HTML ----------
  const DEFAULT_GROUPS = [
    { id: "g1", label: "3 -10 crs" },
    { id: "g2", label: "10-50crs" },
    { id: "g3", label: "50-105crs" },
    { id: "g4", label: "Greater than 105 cr" },
  ];

  function autoAssignSlabGroups(sheet) {
    if (!sheet || !sheet.slabs || sheet.slabs.length === 0) return;
    sheet.slabs.forEach((s) => {
      const f = parseAmt(s.amtFrom);
      const t = parseAmt(s.amtTo);
      if (t != null && t <= 100000000) {
        s.groupId = "g1";
      } else if (t != null && t <= 500000000) {
        s.groupId = "g2";
      } else if (t != null && t <= 1050000000) {
        s.groupId = "g3";
      } else {
        s.groupId = "g4";
      }
    });
  }

  function groupsOrdered(sheet) {
    if (!sheet || !sheet.slabs || sheet.slabs.length === 0) return [];
    autoAssignSlabGroups(sheet);
    const result = [];
    sheet.slabs.forEach((s) => {
      const gDef = DEFAULT_GROUPS.find((g) => g.id === s.groupId) || DEFAULT_GROUPS[3];
      if (result.length > 0 && result[result.length - 1].id === gDef.id) {
        result[result.length - 1].count++;
      } else {
        result.push({ id: gDef.id, label: gDef.label, count: 1 });
      }
    });
    return result;
  }

  function bandStarts(sheet) {
    const set = new Set();
    if (!sheet || !sheet.slabs || sheet.slabs.length === 0) return set;
    autoAssignSlabGroups(sheet);
    sheet.slabs.forEach((s, idx) => {
      if (idx > 0 && s.groupId !== sheet.slabs[idx - 1].groupId) {
        set.add(s.id);
      }
    });
    return set;
  }

  function toISODate(dStr) {
    if (!dStr) return new Date().toISOString().split("T")[0];
    if (/^\d{4}-\d{2}-\d{2}$/.test(dStr)) return dStr;
    const parts = dStr.split("-");
    if (parts.length === 3) {
      const day = parts[0].padStart(2, "0");
      const months = { jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06", jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12" };
      const mKey = parts[1].toLowerCase().slice(0, 3);
      const month = months[mKey] || "09";
      let year = parts[2];
      if (year.length === 2) year = "20" + year;
      return `${year}-${month}-${day}`;
    }
    return new Date().toISOString().split("T")[0];
  }

  function getEffStamp(effDate) {
    const effStr = toISODate(effDate).replace(/-/g, "");
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    return effStr + "_" + hh + mm;
  }

  function fmtDateDisp(isoStr) {

    if (!isoStr) return "";
    const parts = isoStr.split("-");
    if (parts.length === 3) {
      const year = parts[0].slice(2);
      const mIdx = parseInt(parts[1], 10) - 1;
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const monthName = months[mIdx] || "Aug";
      const day = parts[2].padStart(2, "0");
      return `${day} ${monthName} ${year}`;
    }
    return isoStr;
  }

  function gridHTML(sheet, editable, scope = "work") {
    if (!sheet) return "";
    sheet.slabs = sheet.slabs || [];
    sheet.groups = sheet.groups || [];
    sheet.tenures = sheet.tenures || [];
    sheet.rates = sheet.rates || {};
    const bs = bandStarts(sheet);
    const groups = groupsOrdered(sheet);
    const accents = ["#818cf8", "#38bdf8", "#34d399", "#fbbf24"];
    const isModal = (scope === "modal");
    const minWStyle = isModal ? "" : ' style="min-width:' + Math.max(1850, 180 + sheet.slabs.length * 135) + 'px"';
    const wrapCls = "gridwrap" + (isModal ? " modal-gridwrap" : "");
    let h = '<div class="' + wrapCls + '"><table class="grid"' + minWStyle + '><thead>';

    // Row 1: effective date (without group bands)
    h += "<tr>";
    h += '<th class="h-ink" style="text-align:left"><span class="lbl-eff">Effective Date</span></th>';
    h += '<th class="h-ink">' + (editable
      ? '<div style="display:flex;flex-direction:column;align-items:center;gap:3px;"><input type="date" class="h-in date-in" data-role="effDate" data-scope="' + scope + '" value="' + toISODate(sheet.effectiveDate) + '" style="font-weight:600;color:#fff;background:transparent;border:1px solid rgba(255,255,255,.2);border-radius:6px;padding:2px 4px;color-scheme:dark;"><span style="font-size:11px;font-weight:700;color:#38bdf8;">' + esc(fmtDateDisp(toISODate(sheet.effectiveDate))) + '</span></div>'
      : '<span style="font-weight:700;color:#38bdf8;">' + esc(fmtDateDisp(toISODate(sheet.effectiveDate)) || sheet.effectiveDate) + "</span>") + "</th>";
    const remCols = sheet.slabs.length + (editable ? 1 : 0);
    h += '<th class="h-ink" colspan="' + remCols + '"></th>';
    h += "</tr>";

    // Row 2: product + slab labels
    h += "<tr>";
    h += '<th class="h-slab" style="text-align:left"><span style="font-size:12px;font-weight:600">' + esc(sheet.productLabel) + "</span></th>";
    h += '<th class="h-slab"><span style="font-size:11px;color:#cbd5e1">Amt Slab (in Crs)</span></th>';
    sheet.slabs.forEach((s, sIdx) => {
      const cls = "h-slab" + (s.tint ? " h-tint" : "") + (bs.has(s.id) ? " band-div" : "");
      h += '<th class="' + cls + '" data-slabhead="' + s.id + '">';
      if (editable) {
        h += '<div style="display:flex;align-items:center;gap:4px">' +
          '<textarea class="h-in slab-ta" data-role="slabLabel" data-scope="' + scope + '" data-s="' + s.id + '" data-sidx="' + sIdx + '">' + esc(s.label) + '</textarea>' +
          '<button class="iconbtn" title="Insert slab after" data-action="insSlab" data-scope="' + scope + '" data-s="' + s.id + '" style="color:rgba(255,255,255,.7)">+</button>' +
          (sheet.slabs.length > 1 ? '<button class="slab-x" title="Remove slab" data-action="delSlab" data-scope="' + scope + '" data-s="' + s.id + '">x</button>' : "") +
          "</div>";
      } else {
        h += '<span class="slabname">' + esc(s.label) + "</span>";
      }
      h += "</th>";
    });
    if (editable) h += '<th class="h-slab"></th>';
    h += "</tr>";

    // Row 3: Amt From
    h += "<tr>";
    h += '<th class="h-sub"></th><th class="h-sub"><span class="amtlbl">Amt From</span></th>';
    sheet.slabs.forEach((s, sIdx) => {
      const cls = "h-sub" + (s.tint ? " h-tint2" : "") + (bs.has(s.id) ? " band-div" : "");
      const isFirst = (sIdx === 0);
      const formattedVal = fmtINR(parseAmt(s.amtFrom));
      const displayVal = formattedVal || esc(s.amtFrom);
      h += '<th class="' + cls + '">' + (editable
        ? '<input class="h-in amt-in" data-role="amtFrom" data-scope="' + scope + '" data-s="' + s.id + '" data-sidx="' + sIdx + '" value="' + esc(displayVal) + '" ' + (isFirst ? 'readonly style="background:rgba(255,255,255,0.08);color:#94a3b8;cursor:not-allowed;"' : '') + ' title="Raw: ' + esc(s.amtFrom) + '">'
        : '<span class="amt">' + esc(displayVal) + "</span>") + "</th>";
    });
    if (editable) h += '<th class="h-sub"></th>';
    h += "</tr>";

    // Row 4: Amt To
    h += "<tr>";
    h += '<th class="h-sub"></th><th class="h-sub"><span class="amtlbl">Amt To</span></th>';
    sheet.slabs.forEach((s, sIdx) => {
      const cls = "h-sub" + (s.tint ? " h-tint2" : "") + (bs.has(s.id) ? " band-div" : "");
      const isLast = (sIdx === sheet.slabs.length - 1);
      const formattedVal = fmtINR(parseAmt(s.amtTo));
      const displayVal = formattedVal || esc(s.amtTo);
      h += '<th class="' + cls + '">' + (editable
        ? '<input class="h-in amt-in" data-role="amtTo" data-scope="' + scope + '" data-s="' + s.id + '" data-sidx="' + sIdx + '" value="' + esc(displayVal) + '" ' + (isLast ? 'readonly style="background:rgba(255,255,255,0.08);color:#94a3b8;cursor:not-allowed;"' : '') + ' title="Raw: ' + esc(s.amtTo) + '">'
        : '<span class="amt">' + esc(displayVal) + "</span>") + "</th>";
    });
    if (editable) h += '<th class="h-sub"></th>';
    h += "</tr>";

    // Row 5: tenure headers + Callable
    h += "<tr>";
    h += '<th class="h-slab"><span class="tenhdr">Tenure<br>From Month</span></th>';
    h += '<th class="h-slab"><span class="tenhdr">Tenure<br>To Days</span></th>';
    sheet.slabs.forEach((s) => {
      const cls = "h-slab" + (s.tint ? " h-tint" : "") + (bs.has(s.id) ? " band-div" : "");
      h += '<th class="' + cls + '"><span class="callbl">' + esc(s.sub) + "</span></th>";
    });
    if (editable) h += '<th class="h-slab"></th>';
    h += "</tr></thead><tbody>";

    // Body
    sheet.tenures.forEach((t, tIdx) => {
      h += '<tr data-tenrow="' + t.id + '">';
      if (editable) {
        h += '<td class="tencell" data-tencell="' + t.id + '" data-tidx="' + tIdx + '"><input class="ten-in" data-role="tmonth" data-scope="' + scope + '" data-t="' + t.id + '" data-tidx="' + tIdx + '" data-sidx="0" value="' + esc(t.fromMonth) + '"></td>';
        h += '<td class="tencell" data-tencell="' + t.id + '" data-tidx="' + tIdx + '"><input class="ten-in" data-role="tdays" data-scope="' + scope + '" data-t="' + t.id + '" data-tidx="' + tIdx + '" data-sidx="0" value="' + esc(t.toDays) + '"></td>';
      } else {
        h += '<td class="tencell">' + esc(t.fromMonth) + "</td><td class=\"tencell\">" + esc(t.toDays) + "</td>";
      }
      sheet.slabs.forEach((s, sIdx) => {
        const cls = (s.tint ? "tint " : "") + (bs.has(s.id) ? "band-div-l" : "");
        const idxKey = "idx_" + tIdx + "_" + sIdx;
        const legacyKey = rkey(t.id, s.id);
        const v = sheet.rates[idxKey] !== undefined ? sheet.rates[idxKey] : (sheet.rates[legacyKey] || "");
        h += '<td class="' + cls + '" data-cell="' + legacyKey + '" data-tidx="' + tIdx + '" data-sidx="' + sIdx + '">';
        if (editable) {
          h += '<span style="display:inline-flex;align-items:center"><input class="rate-in" data-role="rate" data-scope="' + scope + '" data-t="' + t.id + '" data-s="' + s.id + '" data-tidx="' + tIdx + '" data-sidx="' + sIdx + '" value="' + esc(v) + '"><span class="pct">%</span></span>';
        } else {
          h += esc(v) + "%";
        }
        h += "</td>";
      });
      if (editable) {
        h += '<td class="rowacts">' +
          '<button class="iconbtn" title="Insert row below" data-action="insTen" data-scope="' + scope + '" data-i="' + tIdx + '">+</button>' +
          (sheet.tenures.length > 1 ? '<button class="iconbtn del" title="Delete row" data-action="delTen" data-scope="' + scope + '" data-t="' + t.id + '">x</button>' : "") +
          "</td>";
      }
      h += "</tr>";
    });
    h += "</tbody></table>";

    if (editable) {
      h += '<div class="addbar">' +
        '<button class="addbtn" data-action="addTen" data-scope="' + scope + '">+ Add tenure row</button>' +
        '<button class="addbtn" data-action="addSlab" data-scope="' + scope + '">+ Add amount slab</button></div>';
    }
    h += "</div>";
    return h;
  }

  // ---------- mutations ----------
  function getSheet(scope) {
    return scope === "pub" ? cur().published : cur().working;
  }

  function fmtCrores(val) {
    const n = parseAmt(val);
    if (n == null) return "";
    if (n >= 10000000) {
      const cr = n / 10000000;
      return (Number.isInteger(cr) ? cr : cr.toFixed(2)) + " cr";
    } else if (n >= 100000) {
      const lk = n / 100000;
      return (Number.isInteger(lk) ? lk : lk.toFixed(2)) + " L";
    }
    return fmtINR(n);
  }

  function autoSlabLabel(af, at) {
    const f = parseAmt(af), t = parseAmt(at);
    if (f == null && t == null) return "New slab";
    if (f === 30000000 && t === 32499999) return "3cr to <3.25cr";
    if (f === 32500000 && t === 34999999) return "3.25cr to <3.50cr";
    if (f === 35000000 && t === 100000000) return "3.50cr to 10cr";
    if (f === 100000001 && t === 150000000) return "Above 10 cr to 15 cr";
    if (f === 150000001 && t === 200000000) return "Above 15 cr to 20 cr";
    if (f === 200000001 && t === 250000000) return "Above 20 cr to 25 cr";
    if (f === 250000001 && t === 350000000) return "Above 25 cr to 35 cr";
    if (f === 350000001 && t === 500000000) return "Above 35 cr to 50 cr";
    if (f === 500000001 && t === 700000000) return "Above 50 cr to 70 cr";
    if (f === 700000001 && t === 950000000) return "Above 70 cr to 95 cr";
    if (f === 950000001 && t === 1050000000) return "Above 95 cr to 105 cr";
    if (f === 1050000001 && t === 1500000000) return "Above 105 cr to 150 cr";
    if (f === 150000001 && t === 200000000) return "Above 150 cr to 200 cr";
    if (t == null || t >= 9999999999) return "Above " + fmtCrores(f);

    const prefix = (f > 100000000) ? "Above " : "";
    return prefix + fmtCrores(f) + " to " + fmtCrores(t);
  }

  function rebalanceSlabs(w) {
    if (!w || !w.slabs || w.slabs.length === 0) return;
    w.groups = DEFAULT_GROUPS;
    for (let i = 0; i < w.slabs.length; i++) {
      if (i === 0) {
        w.slabs[0].amtFrom = "30000000";
      }
      if (i > 0) {
        const prevTo = parseAmt(w.slabs[i - 1].amtTo);
        if (prevTo != null && !isNaN(prevTo)) {
          const expectedFrom = prevTo + 1;
          w.slabs[i].amtFrom = String(expectedFrom);
          let currentTo = parseAmt(w.slabs[i].amtTo);
          if (i < w.slabs.length - 1) {
            if (currentTo == null || isNaN(currentTo) || currentTo <= expectedFrom) {
              currentTo = expectedFrom + 10000000;
              w.slabs[i].amtTo = String(currentTo);
            }
          }
        }
      }
      if (i === w.slabs.length - 1) {
        w.slabs[i].amtTo = "99999999999";
      }

      const f = parseAmt(w.slabs[i].amtFrom);
      const t = parseAmt(w.slabs[i].amtTo);
      if (t != null && t <= 100000000) {
        w.slabs[i].groupId = "g1";
      } else if (t != null && t <= 500000000) {
        w.slabs[i].groupId = "g2";
      } else if (t != null && t <= 1050000000) {
        w.slabs[i].groupId = "g3";
      } else {
        w.slabs[i].groupId = "g4";
      }

      w.slabs[i].label = autoSlabLabel(w.slabs[i].amtFrom, w.slabs[i].amtTo);
    }
  }

  function updateSlabDOM(w, scope = "work") {
    w.slabs.forEach((s) => {
      const lblEl = document.querySelector('[data-role="slabLabel"][data-s="' + s.id + '"][data-scope="' + scope + '"]');
      if (lblEl && document.activeElement !== lblEl) {
        lblEl.value = s.label;
      }
      const fromEl = document.querySelector('[data-role="amtFrom"][data-s="' + s.id + '"][data-scope="' + scope + '"]');
      if (fromEl && document.activeElement !== fromEl) {
        const pf = parseAmt(s.amtFrom);
        fromEl.value = pf != null ? fmtINR(pf) : s.amtFrom;
      }
      const toEl = document.querySelector('[data-role="amtTo"][data-s="' + s.id + '"][data-scope="' + scope + '"]');
      if (toEl && document.activeElement !== toEl) {
        const pt = parseAmt(s.amtTo);
        toEl.value = pt != null ? fmtINR(pt) : s.amtTo;
      }
    });
  }

  function rebuildRatesMap(w) {
    const newRates = {};
    w.tenures.forEach((t, tIdx) => {
      w.slabs.forEach((s, sIdx) => {
        const idxKey = "idx_" + tIdx + "_" + sIdx;
        const legacyKey = rkey(t.id, s.id);
        const val = w.rates[legacyKey] !== undefined ? w.rates[legacyKey] : (w.rates[idxKey] || "5.00");
        newRates[idxKey] = val;
        newRates[legacyKey] = val;
      });
    });
    w.rates = newRates;
  }

  function insSlabAfter(slabId, scope = "work") {
    const w = getSheet(scope);
    const idx = w.slabs.findIndex((s) => s.id === slabId);
    if (idx < 0) return;
    const prev = w.slabs[idx];
    const prevFrom = parseAmt(prev.amtFrom) || 30000000;
    const prevTo = parseAmt(prev.amtTo) || prevFrom;

    const halfStep = Math.max(1000000, Math.floor((prevTo - prevFrom) / 2));
    const newPrevTo = (prevTo >= 9999999999) ? (prevFrom + 10000000) : Math.max(prevFrom + 1, prevFrom + halfStep);
    prev.amtTo = String(newPrevTo);

    const newFrom = String(newPrevTo + 1);
    const newTo = String((prevTo >= 9999999999) ? 999999999999 : prevTo);

    const ns = {
      id: nid("s"),
      groupId: prev.groupId,
      label: autoSlabLabel(newFrom, newTo),
      amtFrom: newFrom,
      amtTo: newTo,
      tint: false,
      sub: prev.sub,
      _addedByUser: true,
    };

    w.tenures.forEach((t, tIdx) => {
      const leftKey = "idx_" + tIdx + "_" + idx;
      const leftVal = w.rates[leftKey] !== undefined ? w.rates[leftKey] : (w.rates[rkey(t.id, prev.id)] || "5.00");
      w.rates[rkey(t.id, ns.id)] = leftVal;
    });

    w.slabs.splice(idx + 1, 0, ns);
    rebuildRatesMap(w);
    rebalanceSlabs(w);
    return ns.id;
  }

  function addSlabEnd(scope = "work") {
    const w = getSheet(scope);
    const lastSlab = w.slabs[w.slabs.length - 1];
    const lastIdx = w.slabs.length - 1;
    let newFrom = 1050000001;
    if (lastSlab) {
      const lastFrom = parseAmt(lastSlab.amtFrom) || 100000000;
      const lastTo = parseAmt(lastSlab.amtTo);
      if (lastTo == null || lastTo >= 9999999999) {
        lastSlab.amtTo = String(lastFrom + 50000000);
        newFrom = lastFrom + 50000001;
      } else {
        newFrom = lastTo + 1;
      }
    }
    const newTo = 999999999999;
    const lastGroup = w.groups[w.groups.length - 1] ? w.groups[w.groups.length - 1].id : "g1";
    const sub = w.slabs[0] ? w.slabs[0].sub : LABELS[productKey];
    const ns = {
      id: nid("s"),
      groupId: lastGroup,
      label: autoSlabLabel(String(newFrom), String(newTo)),
      amtFrom: String(newFrom),
      amtTo: String(newTo),
      tint: false,
      sub,
      _addedByUser: true,
    };

    w.tenures.forEach((t, tIdx) => {
      const leftKey = lastSlab ? ("idx_" + tIdx + "_" + lastIdx) : null;
      const leftVal = (lastSlab && w.rates[leftKey] !== undefined) ? w.rates[leftKey] : (lastSlab ? (w.rates[rkey(t.id, lastSlab.id)] || "5.00") : "5.00");
      w.rates[rkey(t.id, ns.id)] = leftVal;
    });

    w.slabs.push(ns);
    rebuildRatesMap(w);
    rebalanceSlabs(w);
    return ns.id;
  }

  function delSlab(slabId, scope = "work") {
    const w = getSheet(scope);
    w.slabs = w.slabs.filter((s) => s.id !== slabId);
    w.tenures.forEach((t) => delete w.rates[rkey(t.id, slabId)]);
    if (w.slabs.length > 0) {
      w.slabs[0].amtFrom = "30000000";
      w.slabs[w.slabs.length - 1].amtTo = "99999999999";
    }
    rebuildRatesMap(w);
    rebalanceSlabs(w);
  }

  function insTenAfter(idx, scope = "work") {
    const w = getSheet(scope);
    const prev = w.tenures[idx];
    const nt = { id: nid("t"), fromMonth: prev ? prev.fromMonth : 0, toDays: prev ? (+prev.toDays + 1) : 0, _addedByUser: true };

    const c = cur();
    const pub = c ? c.published : null;

    w.slabs.forEach((s, sIdx) => {
      let pubVal = getPublishedRate(pub, s, nt, idx + 1, sIdx);
      if (!pubVal || pubVal === "") {
        const upperKey = prev ? ("idx_" + idx + "_" + sIdx) : null;
        pubVal = (prev && w.rates[upperKey] !== undefined) ? w.rates[upperKey] : (prev ? (w.rates[rkey(prev.id, s.id)] || "5.00") : "5.00");
      }
      w.rates[rkey(nt.id, s.id)] = pubVal;
    });

    w.tenures.splice(idx + 1, 0, nt);
    rebuildRatesMap(w);
    return nt.id;
  }

  function addTenEnd(scope = "work") {
    const w = getSheet(scope);
    return insTenAfter(w.tenures.length - 1, scope);
  }

  function delTen(tId, scope = "work") {
    const w = getSheet(scope);
    w.tenures = w.tenures.filter((t) => t.id !== tId);
    w.slabs.forEach((s) => delete w.rates[rkey(tId, s.id)]);
    rebuildRatesMap(w);
  }

  // ---------- render ----------
  function renderTabs() {
    const el = document.getElementById("tabs");
    if (!el) return;
    const activeProducts = (DATA && Object.keys(DATA).length > 0) ? Object.keys(DATA) : PRODUCTS;
    if (DATA && !DATA[productKey] && activeProducts.length > 0) {
      productKey = activeProducts[0];
    }
    el.innerHTML = activeProducts.map((p) => {
      const d = DATA ? DATA[p] : null;
      let diffCount = 0;
      if (d && d.published && d.working) {
        Object.keys(d.working.rates || {}).forEach((k) => {
          if (k.indexOf("idx_") === 0) {
            if ((d.published.rates[k] || "") !== (d.working.rates[k] || "")) {
              diffCount++;
            }
          }
        });
      }
      const labelText = LABELS[p] || (d && d.published ? d.published.productLabel : p);
      return '<button class="tab ' + (p === productKey ? "active" : "") + '" data-tab="' + p + '">' +
        esc(labelText) + (diffCount > 0 ? ' <span style="background:#d97706;color:#fff;font-size:11px;font-weight:700;padding:1px 6px;border-radius:10px;margin-left:6px;" title="' + diffCount + ' user change(s)">' + diffCount + ' draft change(s)</span>' : "") + "</button>";
    }).join("");
    el.querySelectorAll("[data-tab]").forEach((b) => b.addEventListener("click", () => { productKey = b.dataset.tab; saved = false; render(); }));
  }

  function render(focusNewId) {
    const workWrap = document.querySelector("#working .gridwrap") || document.querySelector("#working .table-wrap");
    const pubWrap = document.querySelector("#published .gridwrap") || document.querySelector("#published .table-wrap");
    const workScrollLeft = workWrap ? workWrap.scrollLeft : 0;
    const workScrollTop = workWrap ? workWrap.scrollTop : 0;
    const pubScrollLeft = pubWrap ? pubWrap.scrollLeft : 0;
    const pubScrollTop = pubWrap ? pubWrap.scrollTop : 0;
    const scrollY = window.scrollY || document.documentElement.scrollTop;

    renderTabs();
    const c = cur();
    if (!c) return;
    let pubSheet = c.published || c.working;
    let workSheet = c.working || c.published;
    if (!pubSheet || !workSheet) return;

    if (isNominalMode) {
      pubSheet = convertSheetToNominalMonthly(pubSheet);
      workSheet = convertSheetToNominalMonthly(workSheet);
    }

    const pLabel = LABELS[productKey] || pubSheet.productLabel || "Treasury Rates";
    const modeSuffix = isNominalMode ? " (Monthly Nominal Rates)" : "";

    const pubTitleEl = document.getElementById("pub-title");
    if (pubTitleEl) pubTitleEl.textContent = "Published rates - " + pLabel + modeSuffix;
    const workTitleEl = document.getElementById("work-title");
    if (workTitleEl) workTitleEl.textContent = "Working draft - " + pLabel + modeSuffix;
    const pubNoteEl = document.getElementById("pub-note");
    if (pubNoteEl) pubNoteEl.textContent =
      "The fixed table that replaces the Excel sheet for " + pLabel + ". It changes when an admin saves published rates or publishes the draft below.";

    const btnNominal = document.getElementById("btn-nominal");
    if (btnNominal) {
      if (isNominalMode) {
        btnNominal.style.background = "#d97706";
        btnNominal.style.color = "#ffffff";
        btnNominal.style.borderColor = "#b45309";
        btnNominal.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg> ? Nominal Monthly Mode';
      } else {
        btnNominal.style.background = "";
        btnNominal.style.color = "#b45309";
        btnNominal.style.borderColor = "#d97706";
        btnNominal.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg> Nominal Monthly Rate';
      }
    }

    const pc = document.getElementById("published");
    if (pc) {
      pc.innerHTML = gridHTML(pubSheet, IS_ADMIN && !isNominalMode, "pub");
      if (IS_ADMIN && !isNominalMode) {
        pc.oninput = onInput;
        pc.onchange = onChange;
        pc.onclick = onClick;
        pc.onkeydown = onKeyDown;
      }
    }

    const wc = document.getElementById("working");
    if (wc) {
      wc.innerHTML = gridHTML(workSheet, !isNominalMode, "work");
      if (!isNominalMode) {
        wc.oninput = onInput;
        wc.onchange = onChange;
        wc.onclick = onClick;
        wc.onkeydown = onKeyDown;
      }
    }

    const btnSaveEl = document.getElementById("btn-save");
    if (btnSaveEl) {
      btnSaveEl.textContent = "Download " + LABELS[productKey] + " Rates";
    }
    refresh();

    const newWorkWrap = document.querySelector("#working .gridwrap") || document.querySelector("#working .table-wrap");
    const newPubWrap = document.querySelector("#published .gridwrap") || document.querySelector("#published .table-wrap");
    if (newWorkWrap) {
      newWorkWrap.scrollLeft = workScrollLeft;
      newWorkWrap.scrollTop = workScrollTop;
    }
    if (newPubWrap) {
      newPubWrap.scrollLeft = pubScrollLeft;
      newPubWrap.scrollTop = pubScrollTop;
    }
    window.scrollTo(0, scrollY);

    if (focusNewId) {
      setTimeout(() => {
        const newEl = document.querySelector('[data-tenrow="' + focusNewId + '"]') || document.querySelector('[data-slabhead="' + focusNewId + '"]');
        if (newEl) {
          newEl.scrollIntoView({ block: "nearest", inline: "nearest" });
          const firstInp = newEl.querySelector("input");
          if (firstInp) { firstInp.focus(); if (firstInp.select) firstInp.select(); }
        }
      }, 30);
    }
  }

  function getPublishedRate(pub, workSlab, workTenure, tIdx, sIdx) {
    if (!pub || !pub.rates || !workSlab || !workTenure) return "";
    pub.tenures = pub.tenures || [];
    pub.slabs = pub.slabs || [];

    // 1. Direct ID match first
    let key = rkey(workTenure.id, workSlab.id);
    if (pub.rates[key] !== undefined) return pub.rates[key];

    // 2. Find matching published tenure & slab by ID or amtFrom/label
    const pTen = pub.tenures.find((t) => t.id === workTenure.id || (+t.fromMonth === +workTenure.fromMonth && +t.toDays === +workTenure.toDays)) || (tIdx !== undefined ? pub.tenures[tIdx] : null);

    let pSlab = pub.slabs.find((s) => s.id === workSlab.id);
    if (!pSlab && workSlab.amtFrom) {
      const wFrom = parseAmt(workSlab.amtFrom);
      pSlab = pub.slabs.find((s) => parseAmt(s.amtFrom) === wFrom);
    }
    if (!pSlab && workSlab.label) {
      pSlab = pub.slabs.find((s) => s.label.trim() === workSlab.label.trim());
    }

    if (pTen && pSlab) {
      key = rkey(pTen.id, pSlab.id);
      if (pub.rates[key] !== undefined) return pub.rates[key];
    }

    // 3. Fallback to index key if available
    if (tIdx !== undefined && sIdx !== undefined) {
      const idxKey = "idx_" + tIdx + "_" + sIdx;
      if (pub.rates[idxKey] !== undefined) return pub.rates[idxKey];
    }

    return "";
  }

  function focusTarget(target) {
    if (!target) return;
    target.focus();
    if (target.select) target.select();
    target.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  function onKeyDown(e) {
    const el = e.target, role = el.dataset.role;
    if (!role) return;

    const scope = el.dataset.scope || "work";
    const tIdx = parseInt(el.dataset.tidx, 10);
    const sIdx = parseInt(el.dataset.sidx, 10);
    const sId = el.dataset.s;

    if (e.key === "ArrowDown" || e.key === "Enter") {
      if (e.key === "Enter" && (role === "amtFrom" || role === "amtTo" || role === "slabLabel")) return;
      e.preventDefault();
      let target = null;
      if (role === "slabLabel") {
        target = document.querySelector('[data-role="amtFrom"][data-scope="' + scope + '"][data-s="' + sId + '"]');
      } else if (role === "amtFrom") {
        target = document.querySelector('[data-role="amtTo"][data-scope="' + scope + '"][data-s="' + sId + '"]');
      } else if (role === "amtTo") {
        target = document.querySelector('[data-role="rate"][data-scope="' + scope + '"][data-tidx="0"][data-sidx="' + (isNaN(sIdx) ? 0 : sIdx) + '"]');
      } else if (!isNaN(tIdx)) {
        target = document.querySelector('[data-role="' + role + '"][data-scope="' + scope + '"][data-tidx="' + (tIdx + 1) + '"][data-sidx="' + (isNaN(sIdx) ? 0 : sIdx) + '"]');
      }
      if (target) focusTarget(target);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      let target = null;
      if (!isNaN(tIdx) && tIdx > 0) {
        target = document.querySelector('[data-role="' + role + '"][data-scope="' + scope + '"][data-tidx="' + (tIdx - 1) + '"][data-sidx="' + (isNaN(sIdx) ? 0 : sIdx) + '"]');
      } else if (!isNaN(tIdx) && tIdx === 0) {
        target = document.querySelector('[data-role="amtTo"][data-scope="' + scope + '"][data-sidx="' + (isNaN(sIdx) ? 0 : sIdx) + '"]');
      } else if (role === "amtTo") {
        target = document.querySelector('[data-role="amtFrom"][data-scope="' + scope + '"][data-s="' + sId + '"]');
      } else if (role === "amtFrom") {
        target = document.querySelector('[data-role="slabLabel"][data-scope="' + scope + '"][data-s="' + sId + '"]');
      } else if (role === "slabLabel" || role === "effDate") {
        window.scrollTo({ top: 0, behavior: "smooth" });
      }
      if (target) focusTarget(target);
    } else if (e.key === "ArrowRight") {
      if (el.selectionStart === el.value.length || (el.selectionStart === 0 && el.selectionEnd === el.value.length)) {
        e.preventDefault();
        let target = null;
        if (!isNaN(sIdx)) {
          target = document.querySelector('[data-role="' + role + '"][data-scope="' + scope + '"][data-tidx="' + tIdx + '"][data-sidx="' + (sIdx + 1) + '"]');
        } else if (sId) {
          const curSlabIdx = getSheet(scope).slabs.findIndex((s) => s.id === sId);
          if (curSlabIdx >= 0) {
            target = document.querySelector('[data-role="' + role + '"][data-scope="' + scope + '"][data-sidx="' + (curSlabIdx + 1) + '"]');
          }
        }
        if (target) focusTarget(target);
      }
    } else if (e.key === "ArrowLeft") {
      if (el.selectionStart === 0) {
        e.preventDefault();
        let target = null;
        if (!isNaN(sIdx) && sIdx > 0) {
          target = document.querySelector('[data-role="' + role + '"][data-scope="' + scope + '"][data-tidx="' + tIdx + '"][data-sidx="' + (sIdx - 1) + '"]');
        } else if (sId) {
          const curSlabIdx = getSheet(scope).slabs.findIndex((s) => s.id === sId);
          if (curSlabIdx > 0) {
            target = document.querySelector('[data-role="' + role + '"][data-scope="' + scope + '"][data-sidx="' + (curSlabIdx - 1) + '"]');
          }
        }
        if (target) focusTarget(target);
      }
    }
  }

  function onInput(e) {
    const el = e.target, role = el.dataset.role, scope = el.dataset.scope || "work";
    if (!role) return;
    const w = getSheet(scope);
    if (role === "effDate") {
      const formatted = fmtDateDisp(el.value);
      PRODUCTS.forEach((p) => {
        if (DATA[p] && DATA[p].working) {
          DATA[p].working.effectiveDate = el.value;
        }
      });
      const spanEl = el.parentElement ? el.parentElement.querySelector("span") : null;
      if (spanEl) spanEl.textContent = formatted;
      saved = false;
      refresh();
      return;
    }
    else if (role === "rate") {
      el.value = el.value.replace(/[^\d.]/g, "");
      const tIdx = el.dataset.tidx;
      const sIdx = el.dataset.sidx;
      if (tIdx !== undefined && sIdx !== undefined) {
        w.rates["idx_" + tIdx + "_" + sIdx] = el.value;
      }
      const rKey = rkey(el.dataset.t, el.dataset.s);
      w.rates[rKey] = el.value;
      w._userEditedRates = w._userEditedRates || new Set();
      w._userEditedRates.add(rKey);
      if (tIdx !== undefined && sIdx !== undefined) {
        w._userEditedRates.add("idx_" + tIdx + "_" + sIdx);
      }
      saved = false;
      refresh();
      return;
    }
    else if (role === "tmonth") {
      el.value = el.value.replace(/[^\d]/g, "");
      setTen(w, el.dataset.t, "fromMonth", el.value);
      saved = false;
      refresh();
      return;
    }
    else if (role === "tdays") {
      el.value = el.value.replace(/[^\d]/g, "");
      setTen(w, el.dataset.t, "toDays", el.value);
      saved = false;
      refresh();
      return;
    }
    else if (role === "slabLabel") {
      setSlab(w, el.dataset.s, "label", el.value);
      saved = false;
      refresh();
      return;
    }
    else if (role === "amtFrom" || role === "amtTo") {
      const cleanVal = el.value.replace(/[^\d]/g, "");
      setSlab(w, el.dataset.s, role, cleanVal);
      rebalanceSlabs(w);
      updateSlabDOM(w, scope);
      saved = false;
      refresh();
      return;
    }
    saved = false;
    refresh();
  }

  function onChange(e) {
    const el = e.target, role = e.target.dataset.role;
    if (role === "amtFrom" || role === "amtTo" || role === "slabLabel" || role === "effDate") {
      render();
    }
  }
  function setTen(w, id, k, v) { const t = w.tenures.find((x) => x.id === id); if (t) { t[k] = v; t._userEdited = true; } }
  function setSlab(w, id, k, v) { const s = w.slabs.find((x) => x.id === id); if (s) { s[k] = v; s._userEdited = true; } }

  function onClick(e) {
    const b = e.target.closest("[data-action]"); if (!b) return;
    const a = b.dataset.action, scope = b.dataset.scope || "work";
    let focusId = null;
    if (a === "insSlab") focusId = insSlabAfter(b.dataset.s, scope);
    else if (a === "delSlab") delSlab(b.dataset.s, scope);
    else if (a === "addSlab") focusId = addSlabEnd(scope);
    else if (a === "insTen") focusId = insTenAfter(+b.dataset.i, scope);
    else if (a === "delTen") delTen(b.dataset.t, scope);
    else if (a === "addTen") focusId = addTenEnd(scope);
    else return;
    saved = false;
    render(focusId);
  }

  function cleanRateVal(v) {
    if (v == null || v === "") return "";
    const s = String(v).replace("%", "").trim();
    const n = parseFloat(s);
    return isNaN(n) ? s : String(n);
  }

  function highlightSheetContainer(container, sheet, pub) {
    if (!container || !sheet) return 0;
    pub = pub || {};
    pub.tenures = pub.tenures || [];
    pub.slabs = pub.slabs || [];
    pub.rates = pub.rates || {};
    const workTenures = sheet.tenures || [];
    const workSlabs = sheet.slabs || [];
    const workRates = sheet.rates || {};

    let changedCount = 0;

    // 1. Highlight rate cells
    container.querySelectorAll("[data-cell]").forEach((td) => {
      const legacyKey = td.dataset.cell;
      const tIdx = parseInt(td.dataset.tidx, 10);
      const sIdx = parseInt(td.dataset.sidx, 10);

      const workTenure = workTenures[tIdx];
      const workSlab = workSlabs[sIdx];
      const pubVal = getPublishedRate(pub, workSlab, workTenure, tIdx, sIdx);
      const workVal = workRates[legacyKey] !== undefined ? workRates[legacyKey] : (workRates["idx_" + tIdx + "_" + sIdx] || "");

      const isNewTenureRow = !!(workTenure && workTenure._addedByUser);
      const isNewSlabCol = !!(workSlab && workSlab._addedByUser);

      const cWork = cleanRateVal(workVal);
      const cPub = cleanRateVal(pubVal);

      const rateValChanged = (cWork !== cPub && cWork !== "");
      const changed = isNewTenureRow || isNewSlabCol || rateValChanged;

      if (changed) {
        changedCount++;
        td.classList.add("changed");
        td.title = "User changed rate (Published rate was: " + pubVal + "%)";
        const inp = td.querySelector("input");
        if (inp) {
          inp.classList.add("changed");
          inp.title = "Published rate was: " + pubVal + "%";
        }
      } else {
        td.classList.remove("changed");
        td.removeAttribute("title");
        const inp = td.querySelector("input");
        if (inp) inp.classList.remove("changed");
      }
    });

    // 2. Highlight tenure rows & cells
    container.querySelectorAll("tr[data-tenrow]").forEach((tr, tIdx) => {
      const id = tr.dataset.tenrow;
      const w = workTenures.find((x) => x.id === id) || workTenures[tIdx];
      if (!w) return;
      const p = pub.tenures ? (pub.tenures.find((x) => x.id === id || (+x.fromMonth === +w.fromMonth && +x.toDays === +w.toDays)) || pub.tenures[tIdx]) : null;
      const isNewTenureRow = !!w._addedByUser;
      const tenureChanged = isNewTenureRow || (p ? (+w.fromMonth !== +p.fromMonth || +w.toDays !== +p.toDays) : false);

      tr.classList.toggle("row-new-highlight", isNewTenureRow);

      tr.querySelectorAll("[data-tencell]").forEach((td) => {
        if (tenureChanged) changedCount++;
        td.classList.toggle("changed", tenureChanged);
        const inp = td.querySelector("input");
        if (inp) inp.classList.toggle("changed", tenureChanged);
      });

      if (isNewTenureRow) {
        tr.querySelectorAll("[data-cell]").forEach((td) => {
          changedCount++;
          td.classList.add("changed");
          td.title = "Newly added tenure row";
          const inp = td.querySelector("input");
          if (inp) {
            inp.classList.add("changed");
            inp.title = "Newly added tenure row";
          }
        });
      }
    });

    // 3. Highlight Slab headers (slabLabel, amtFrom, amtTo) and slab columns if new or modified
    workSlabs.forEach((wSlab, sIdx) => {
      const wFrom = parseAmt(wSlab.amtFrom);
      const wTo = parseAmt(wSlab.amtTo);

      const pSlab = pub.slabs ? (pub.slabs.find((x) => x.id === wSlab.id) || (wFrom !== null ? pub.slabs.find((x) => parseAmt(x.amtFrom) === wFrom) : null) || pub.slabs[sIdx]) : null;

      const pFrom = pSlab ? parseAmt(pSlab.amtFrom) : null;
      const pTo = pSlab ? parseAmt(pSlab.amtTo) : null;

      const isNewSlab = !pSlab || !!wSlab._addedByUser;
      const labelChanged = isNewSlab || ((wSlab.label || "").trim() !== (pSlab ? (pSlab.label || "").trim() : ""));
      const amtFromChanged = isNewSlab || (wFrom !== pFrom);
      const amtToChanged = isNewSlab || (wTo !== pTo);

      // Slab label header
      const labelEl = container.querySelector("[data-role='slabLabel'][data-s='" + wSlab.id + "']");
      if (labelEl) {
        labelEl.classList.toggle("changed", labelChanged);
        if (labelEl.closest("th")) labelEl.closest("th").classList.toggle("changed", labelChanged);
        if (labelChanged) changedCount++;
      }

      // AmtFrom header
      const fromEl = container.querySelector("[data-role='amtFrom'][data-s='" + wSlab.id + "']");
      if (fromEl) {
        fromEl.classList.toggle("changed", amtFromChanged);
        if (fromEl.closest("th")) fromEl.closest("th").classList.toggle("changed", amtFromChanged);
        if (amtFromChanged) changedCount++;
      }

      // AmtTo header
      const toEl = container.querySelector("[data-role='amtTo'][data-s='" + wSlab.id + "']");
      if (toEl) {
        toEl.classList.toggle("changed", amtToChanged);
        if (toEl.closest("th")) toEl.closest("th").classList.toggle("changed", amtToChanged);
        if (amtToChanged) changedCount++;
      }

      // If newly added slab column, highlight all rate cells in this column
      if (isNewSlab) {
        container.querySelectorAll("[data-cell][data-sidx='" + sIdx + "']").forEach((td) => {
          changedCount++;
          td.classList.add("changed");
          td.title = "Newly added slab column";
          const inp = td.querySelector("input");
          if (inp) {
            inp.classList.add("changed");
            inp.title = "Newly added slab column";
          }
        });
      }
    });

    return changedCount;
  }

  function refresh() {
    const c = cur();
    if (!c || !c.working) return;
    const v = validate(c.working);
    const pub = c.published || {};
    const dirty = JSON.stringify(pub) !== JSON.stringify(c.working);

    const wc = document.getElementById("working");
    const changedCount = wc ? highlightSheetContainer(wc, c.working, pub) : 0;

    const st = document.getElementById("status");
    if (st) {
      if (saved) {
        st.className = "status saved";
        st.textContent = "In sync with published";
      } else if (dirty) {
        st.className = "status dirty";
        st.innerHTML = changedCount > 0 ? changedCount + " user change(s) pending review (highlighted)" : "Unsaved changes - Pending review";
      } else {
        st.className = "status sync";
        st.textContent = "In sync with published";
      }
    }

    const eb = document.getElementById("errbox");
    if (v.errors.length) {
      eb.hidden = false;
      eb.innerHTML = '<div class="err-h">! ' + v.errors.length + " issue" + (v.errors.length > 1 ? "s" : "") + " to fix before saving</div><ul>" +
        v.errors.slice(0, 5).map((x) => "<li>" + esc(x) + "</li>").join("") +
        (v.errors.length > 5 ? "<li>+" + (v.errors.length - 5) + " more...</li>" : "") + "</ul>";
    } else { eb.hidden = true; eb.innerHTML = ""; }

    if (wc) {
      wc.querySelectorAll("[data-slabhead]").forEach((th) => th.classList.toggle("invalid", v.badSlabs.has(th.dataset.slabhead)));
      wc.querySelectorAll("[data-role='amtFrom'], [data-role='amtTo']").forEach((inp) => {
        const isBad = v.badSlabs.has(inp.dataset.s);
        inp.classList.toggle("invalid", isBad);
        if (inp.closest("th")) inp.closest("th").classList.toggle("invalid", isBad);
      });
      wc.querySelectorAll("[data-tenrow]").forEach((tr) => tr.classList.toggle("invalid", v.badTenures.has(tr.dataset.tenrow)));
      wc.querySelectorAll("[data-role='tmonth'], [data-role='tdays']").forEach((inp) => {
        const isBad = v.badTenures.has(inp.dataset.t);
        inp.classList.toggle("invalid", isBad);
      });
      wc.querySelectorAll("[data-cell]").forEach((td) => {
        const key = td.dataset.cell;
        const idxKey = "idx_" + td.dataset.tidx + "_" + td.dataset.sidx;
        const isBad = !!(v.badRates && (v.badRates.has(key) || v.badRates.has(idxKey)));
        td.classList.toggle("invalid", isBad);
        const inp = td.querySelector("input");
        if (inp) inp.classList.toggle("invalid", isBad);
      });
    }

    const vPub = validate(c.published);
    const pc = document.getElementById("published");
    if (pc) {
      pc.querySelectorAll("[data-slabhead]").forEach((th) => th.classList.toggle("invalid", vPub.badSlabs.has(th.dataset.slabhead)));
      pc.querySelectorAll("[data-tenrow]").forEach((tr) => tr.classList.toggle("invalid", vPub.badTenures.has(tr.dataset.tenrow)));
    }
    const bSavePub = document.getElementById("btn-save-pub");
    if (bSavePub) bSavePub.disabled = !IS_ADMIN || vPub.errors.length > 0;

    const bSave = document.getElementById("btn-save");
    if (bSave) bSave.disabled = !IS_ADMIN || v.errors.length > 0;
    const bDraft = document.getElementById("btn-draft");
    if (bDraft) bDraft.disabled = v.errors.length > 0;
    const bReset = document.getElementById("btn-reset");
    if (bReset) bReset.disabled = !dirty;
  }

  // ---------- server calls ----------
  async function post(url, body) {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-CSRFToken": csrftoken },
      body: JSON.stringify(body),
    });
    let data = {};
    try { data = await res.json(); } catch (e) { }
    return { ok: res.ok, status: res.status, data };
  }

  async function saveDraft() {
    const v = validate(cur().working);
    if (v.errors.length) { refresh(); alert("Please fix validation issues before saving draft."); return; }
    const r = await post("/" + productKey + "/draft/", { sheet: cur().working });
    if (r.ok && r.data.ok) {
      if (r.data.working) cur().working = r.data.working;
      saved = true;
      render();
      const st = document.getElementById("status");
      if (st) {
        st.className = "status saved";
        st.textContent = "Draft saved & sent to Admin for review!";
      }
    } else {
      const errs = (r.data && r.data.errors) || ["Could not save draft."];
      alert(errs.join("\n"));
    }
  }

  async function savePublished() {
    if (!IS_ADMIN) return;
    const v = validate(cur().published);
    if (v.errors.length) { alert("Please fix validation issues in Published rates:\n" + v.errors.join("\n")); return; }
    const r = await post("/" + productKey + "/save-published/", { sheet: cur().published });
    if (r.ok && r.data.ok) {
      if (r.data.published) cur().published = r.data.published;
      render();
      flash("Published rates saved successfully");
    } else {
      const errs = (r.data && r.data.errors) || ["Could not save published rates."];
      alert(errs.join("\n"));
    }
  }

    async function publish() {
    const v = validate(cur().working);
    if (v.errors.length) { refresh(); alert("Please fix validation issues before downloading."); return; }

    const r = await post("/" + productKey + "/publish/", { sheet: cur().working });
    if (r.ok && r.data.ok) {
      // 1. Download 1st file: Quarterly Flat File CSV
      download(r.data.filename, r.data.content);

      // 2. Download 2nd file: Nominal Monthly Flat File CSV
      setTimeout(() => {
        const nomFilename = r.data.nom_filename || (LABELS[productKey].replace(/ /g, "_").replace(/\//g, "_") + "_Nominal_Monthly_Rates_" + getEffStamp(cur().working ? cur().working.effectiveDate : "") + ".csv");
        const nomContent = r.data.nom_content || exportFlatCSV(convertSheetToNominalMonthly(cur().working));
        download(nomFilename, nomContent);
        flash("Downloaded 2 CSV Files: Quarterly Rates CSV and Nominal Monthly Rates CSV");
      }, 350);
    } else {
      const errs = (r.data && r.data.errors) || ["Download failed."];
      alert(errs.join("\n"));
    }
  }
 
  function download(filename, content) {
    const blob = new Blob([content], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
  }
  function flash(msg) {
    const st = document.getElementById("status");
    if (st) {
      st.className = "status saved"; st.textContent = msg;
    }
    setTimeout(refresh, 1500);
  }

  const btnSavePub = document.getElementById("btn-save-pub");
  if (btnSavePub) btnSavePub.addEventListener("click", savePublished);

  const btnUploadPub = document.getElementById("btn-upload-pub");
  const pubFileInput = document.getElementById("pub-file-input");

  if (btnUploadPub && pubFileInput) {
    btnUploadPub.addEventListener("click", () => pubFileInput.click());
    pubFileInput.addEventListener("change", async () => {
      const file = pubFileInput.files[0];
      if (!file) return;
      const formData = new FormData();
      formData.append("file", file);

      const res = await fetch("/" + productKey + "/upload-published/", {
        method: "POST",
        headers: { "X-CSRFToken": csrftoken },
        body: formData,
      });

      let data = {};
      try { data = await res.json(); } catch (e) { }
      if (res.ok && data.ok) {
        DATA[productKey].published = data.published;
        DATA[productKey].working = data.working;
        pubFileInput.value = "";
        saved = true;
        render();
        flash("Published rates updated from CSV");
      } else {
        const errs = (data && data.errors) || ["File upload failed."];
        alert(errs.join("\n"));
        pubFileInput.value = "";
      }
    });
  }

  async function resetWorkingDraft() {
    cur().working = deep(cur().published);
    saved = true;
    render();
    const st = document.getElementById("status");
    if (st) {
      st.className = "status saved";
      st.textContent = "Reset complete - Restored to original Published Rates";
    }
    try {
      const r = await post("/" + productKey + "/draft/", { sheet: cur().working });
      if (r.ok && r.data && r.data.working) {
        cur().working = r.data.working;
        render();
        const st2 = document.getElementById("status");
        if (st2) {
          st2.textContent = "Reset complete - Restored to original Published Rates";
        }
      }
    } catch (e) { }
  }

  function exportGridCSV(d) {
    const effDate = fmtDateDisp(toISODate(d.effectiveDate)) || d.effectiveDate || "";
    const tenures = d.tenures || [];
    const slabs = d.slabs || [];
    const rates = d.rates || {};

    const csvRows = [];

    function q(val) {
      if (val == null) return '""';
      const s = String(val).replace(/"/g, '""');
      return `"${s}"`;
    }

    // Row 1: Effective Date + Groups
    const r1 = [q("Effective Date"), q(effDate)];
    slabs.forEach((s) => {
      const g = (d.groups || []).find((x) => x.id === s.groupId);
      r1.push(q(g ? g.label : ""));
    });
    csvRows.push(r1.join(","));

    // Row 2: Product + Slab Labels
    const r2 = [q(d.productLabel || "Working Draft"), q("Amt Slab (in Crs)")];
    slabs.forEach((s) => r2.push(q(s.label || "")));
    csvRows.push(r2.join(","));

    // Row 3: Amt From
    const r3 = [q("Amount Range (INR)"), q("Amt From")];
    slabs.forEach((s) => r3.push(q(s.amtFrom || "")));
    csvRows.push(r3.join(","));

    // Row 4: Amt To
    const r4 = [q("Amount Range (INR)"), q("Amt To")];
    slabs.forEach((s) => r4.push(q(s.amtTo || "")));
    csvRows.push(r4.join(","));

    // Row 5: Product Sub Type
    const r5 = [q("Tenure"), q("Product Type")];
    slabs.forEach((s) => r5.push(q(s.sub || "Callable")));
    csvRows.push(r5.join(","));

    // Header Row: Tenure From Month, Tenure To Days
    const rHeader = [q("Tenure From Month"), q("Tenure To Days")];
    slabs.forEach((s) => rHeader.push(q(s.label || "")));
    csvRows.push(rHeader.join(","));

    // Data Rows
    tenures.forEach((t, tIdx) => {
      const row = [q(t.fromMonth), q(t.toDays)];
      slabs.forEach((s, sIdx) => {
        const idxKey = "idx_" + tIdx + "_" + sIdx;
        const legacyKey = rkey(t.id, s.id);
        const val = rates[idxKey] !== undefined ? rates[idxKey] : (rates[legacyKey] || "");
        row.push(q(val ? val + "%" : ""));
      });
      csvRows.push(row.join(","));
    });

    return csvRows.join("\n");
  }



    async function downloadWorkingDraft() {
    const v = validate(cur().working);
    if (v.errors.length) { refresh(); alert("Please fix validation issues before downloading."); return; }

    const c = cur();
    let sheet = c.working;
    if (isNominalMode) {
      sheet = convertSheetToNominalMonthly(sheet);
    }
    const csvContent = exportGridCSV(sheet);
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    const modeSuffix = isNominalMode ? "_Nominal_Monthly_Draft_" : "_Working_Draft_Table_";
    const filename = LABELS[productKey].replace(/[\s/]+/g, "_") + modeSuffix + dateStr + ".csv";
    download(filename, csvContent);
    flash((isNominalMode ? "Nominal Monthly draft" : "Working draft") + " CSV downloaded successfully");
  }

  const btnDownloadDraft = document.getElementById("btn-download-draft");
  if (btnDownloadDraft) btnDownloadDraft.addEventListener("click", downloadWorkingDraft);

  const btnReset = document.getElementById("btn-reset");
  if (btnReset) btnReset.addEventListener("click", resetWorkingDraft);

  const btnDraft = document.getElementById("btn-draft");
  if (btnDraft) btnDraft.addEventListener("click", saveDraft);

  const btnSave = document.getElementById("btn-save");
  if (btnSave) btnSave.addEventListener("click", publish);

 
  async function pollRates() {
    const active = document.activeElement;
    if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA")) return;

    try {
      const res = await fetch("/api/rates/");
      if (res.ok) {
        const json = await res.json();
        if (json.ok && json.data) {
          let updated = false;
          PRODUCTS.forEach((p) => {
            if (json.data[p]) {
              const isDirty = JSON.stringify(DATA[p].working) !== JSON.stringify(DATA[p].published);
              if (!isDirty) {
                const newWorkStr = JSON.stringify(json.data[p].working);
                const oldWorkStr = JSON.stringify(DATA[p].working);
                if (newWorkStr !== oldWorkStr) {
                  DATA[p].working = json.data[p].working;
                  updated = true;
                }
              }
              const newPubStr = JSON.stringify(json.data[p].published);
              const oldPubStr = JSON.stringify(DATA[p].published);
              if (newPubStr !== oldPubStr) {
                DATA[p].published = json.data[p].published;
                updated = true;
              }
            }
          });
          if (updated) {
            render();
          }
        }
      }
    } catch (e) { }
  }

  // ---------- Single View Modal Logic ----------
  let modalScale = 1.0;

  function setModalScale(scale) {
    modalScale = Math.max(0.3, Math.min(2.0, scale));
    const gridWrap = document.getElementById("fullview-grid-wrap");
    const zoomText = document.getElementById("fullview-zoom-val");
    if (gridWrap) {
      gridWrap.style.transform = "scale(" + modalScale + ")";
    }
    if (zoomText) {
      zoomText.textContent = Math.round(modalScale * 100) + "%";
    }
  }

  function autoFitModalScale() {
    const body = document.getElementById("fullview-body");
    const wrap = document.getElementById("fullview-grid-wrap");
    if (!body || !wrap) {
      setModalScale(1.0);
      return;
    }
    wrap.style.transform = "scale(1.0)";
    const bodyW = body.clientWidth - 32;
    const bodyH = body.clientHeight - 32;
    const wrapW = wrap.scrollWidth || wrap.offsetWidth;
    const wrapH = wrap.scrollHeight || wrap.offsetHeight;
    if (bodyW > 0 && bodyH > 0 && wrapW > 0 && wrapH > 0) {
      const scaleX = bodyW / wrapW;
      const scaleY = bodyH / wrapH;
      const fitScale = Math.min(1.0, scaleX, scaleY);
      setModalScale(Math.max(0.35, fitScale));
    } else {
      setModalScale(1.0);
    }
  }

  function openFullViewModal() {
    const modal = document.getElementById("fullview-modal");
    const c = cur();
    if (!modal || !c) return;

    const sheet = c.working || c.published;
    const pLabel = LABELS[productKey] || sheet.productLabel || "Treasury Rates";

    const titleEl = document.getElementById("fullview-title");
    if (titleEl) titleEl.textContent = pLabel + " Rate Card";

    const subTitleEl = document.getElementById("fullview-subtitle");
    if (subTitleEl) subTitleEl.textContent = "Effective Date: " + (fmtDateDisp(toISODate(sheet.effectiveDate)) || sheet.effectiveDate || "Current");

    const badgeEl = document.getElementById("fullview-badge");
    if (badgeEl) {
      badgeEl.textContent = "SINGLE SCREEN VIEW";
      badgeEl.style.background = "";
      badgeEl.style.color = "";
      badgeEl.style.borderColor = "";
    }

    const btnDl = document.getElementById("btn-fullview-download");
    if (btnDl) btnDl.style.display = "none";

    const gridWrap = document.getElementById("fullview-grid-wrap");
    if (gridWrap) {
      gridWrap.innerHTML = gridHTML(sheet, false, "modal");
      highlightSheetContainer(gridWrap, sheet, c.published);
    }

    modal.hidden = false;
    document.body.style.overflow = "hidden";

    setTimeout(autoFitModalScale, 50);
  }

  function closeFullViewModal() {
    const modal = document.getElementById("fullview-modal");
    if (modal) modal.hidden = true;
    document.body.style.overflow = "";
  }

  // ---------- Nominal Monthly Rate Logic ----------
  function toNominalMonthly(r) {
    const num = parseFloat(r);
    if (isNaN(num) || num <= 0) return r;
    // Formula: Convert Quarterly Compounded Rate R (%) to Monthly Compounding Rate Rm (%):
    // Rm = 12 * ((1 + R/400)^(1/3) - 1) * 100
    const monthly = 12 * (Math.pow(1 + num / 400, 1 / 3) - 1) * 100;
    return monthly.toFixed(2);
  }

  function convertSheetToNominalMonthly(sheet) {
    if (!sheet) return sheet;
    const newSheet = JSON.parse(JSON.stringify(sheet));
    newSheet.isNominalMonthly = true;
    const rates = newSheet.rates || {};
    const newRates = {};
    Object.keys(rates).forEach((key) => {
      if (key.startsWith("_")) {
        newRates[key] = rates[key];
      } else {
        const origVal = rates[key];
        if (origVal !== "" && origVal != null) {
          newRates[key] = toNominalMonthly(origVal);
        } else {
          newRates[key] = origVal;
        }
      }
    });
    newSheet.rates = newRates;
    return newSheet;
  }

  function openNominalMonthlyModal() {
    const modal = document.getElementById("fullview-modal");
    const c = cur();
    if (!modal || !c) return;

    const sheet = c.working || c.published;
    const nomSheet = convertSheetToNominalMonthly(sheet);
    const pLabel = LABELS[productKey] || sheet.productLabel || "Treasury Rates";

    const titleEl = document.getElementById("fullview-title");
    if (titleEl) titleEl.textContent = pLabel + " — Monthly Nominal Rate Card";

    const subTitleEl = document.getElementById("fullview-subtitle");
    if (subTitleEl) subTitleEl.textContent = "Effective Date: " + (fmtDateDisp(toISODate(sheet.effectiveDate)) || sheet.effectiveDate || "Current") + " (Compounded Monthly Equivalent)";

    const badgeEl = document.getElementById("fullview-badge");
    if (badgeEl) {
      badgeEl.textContent = "NOMINAL MONTHLY RATE";
      badgeEl.style.background = "#d97706";
      badgeEl.style.color = "#ffffff";
      badgeEl.style.borderColor = "#d97706";
    }

    const btnDl = document.getElementById("btn-fullview-download");
    if (btnDl) btnDl.style.display = "inline-block";

    const gridWrap = document.getElementById("fullview-grid-wrap");
    if (gridWrap) {
      gridWrap.innerHTML = gridHTML(nomSheet, false, "modal");
    }

    modal.hidden = false;
    document.body.style.overflow = "hidden";

    setTimeout(autoFitModalScale, 50);
  }

  function exportFlatCSV(sheet) {
    if (!sheet) return "";
    const headers = [
      "EffectiveDate", "StartMonths", "StartDays", "AmountFrom", "AmountTo",
      "InterestIndexRate", "InterestVarianceRate", "PenaltyIndexRate", "PenaltyVarianceRate"
    ];

    const tenures = (sheet.tenures || []).slice().sort((a, b) => {
      const ma = parseInt(a.fromMonth || 0, 10), mb = parseInt(b.fromMonth || 0, 10);
      if (ma !== mb) return ma - mb;
      return parseInt(a.toDays || 0, 10) - parseInt(b.toDays || 0, 10);
    });

    const slabs = (sheet.slabs || []).slice().sort((a, b) => {
      const fa = parseFloat(String(a.amtFrom || "").replace(/[^\d.]/g, "")) || 0;
      const fb = parseFloat(String(b.amtFrom || "").replace(/[^\d.]/g, "")) || 0;
      return fa - fb;
    });

    const rates = sheet.rates || {};
    const date = sheet.effectiveDate || "06-Aug-26";
    const isNom = !!sheet.isNominalMonthly;

    const lines = [headers.join(",")];

    tenures.forEach((t, tIdx) => {
      slabs.forEach((s, sIdx) => {
        const idxKey = "idx_" + tIdx + "_" + sIdx;
        const legacyKey = rkey(t.id, s.id);
        let val = rates[idxKey] !== undefined ? rates[idxKey] : (rates[legacyKey] || "");
        if (val !== "" && val != null) {
        //   if (isNom) {
        //     val = toNominalMonthly(val);
        //   }
          if (!String(val).endsWith("%")) val = val + "%";
        } else {
          val = "";
        }

        const amtFrom = String(s.amtFrom || "").replace(/[^\d.]/g, "");
        const amtTo = String(s.amtTo || "").replace(/[^\d.]/g, "");

        lines.push([
          date,
          t.fromMonth || "0",
          t.toDays || "0",
          amtFrom,
          amtTo,
          val,
          "0",
          "1%",
          "0"
        ].join(","));
      });
    });

    return lines.join("\n");
  }

  function downloadNominalMonthlyCSV() {
    const c = cur();
    if (!c) return;
    const sheet = c.working || c.published;
    const nomSheet = convertSheetToNominalMonthly(sheet);
    const csvContent = exportFlatCSV(nomSheet);
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    const filename = LABELS[productKey].replace(/[\s/]+/g, "_") + "_Nominal_Monthly_Rates_" + dateStr + ".csv";
    download(filename, csvContent);
    flash("Nominal Monthly Rates flat file (OS Spin format) CSV downloaded successfully");
  }

  // ---------- Rate Card History Modal Logic ----------
  async function openHistoryModal() {
    const modal = document.getElementById("history-modal");
    const wrap = document.getElementById("history-list-wrap");
    const titleEl = document.getElementById("history-modal-title");
    if (!modal || !wrap) return;

    const pLabel = LABELS[productKey] || productKey;
    if (titleEl) titleEl.textContent = pLabel + " — Rate Card History";

    wrap.innerHTML = '<div style="text-align:center;padding:24px;color:#64748b;font-weight:500;">Loading database rate history...</div>';
    modal.hidden = false;
    document.body.style.overflow = "hidden";

    try {
      const res = await fetch("/api/" + productKey + "/history/", {
        headers: { "X-Requested-With": "XMLHttpRequest" }
      });
      const data = await res.json();

      if (!data.ok || !data.history || data.history.length === 0) {
        wrap.innerHTML = '<div style="text-align:center;padding:24px;color:#94a3b8;font-size:14px;">No historical rate cards found in database yet.</div>';
        return;
      }

      let html = '<div style="display:flex;flex-direction:column;gap:10px;">';
      data.history.forEach((h) => {
        const dateDisp = fmtDateDisp(toISODate(h.effective_date)) || h.effective_date;
        html += '<div class="history-item" data-hid="' + h.id + '" style="display:flex;align-items:center;justify-content:space-between;padding:12px 16px;background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;cursor:pointer;transition:all .15s ease;" onmouseover="this.style.borderColor=\'#6366f1\';this.style.background=\'#f0fdf4\';" onmouseout="this.style.borderColor=\'#e2e8f0\';this.style.background=\'#ffffff\';">';
        html += '<div>';
        html += '<div style="font-weight:700;color:#1e293b;font-size:14px;display:flex;align-items:center;gap:6px;">' +
          '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#6366f1" stroke-width="2.5"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>' +
          'Effective: ' + esc(dateDisp) + '</div>';
        html += '<div style="font-size:11px;color:#64748b;margin-top:2px;">Last updated: ' + esc(h.updated_at) + '</div>';
        html += '</div>';
        html += '<button class="btn primary" style="padding:6px 12px;font-size:12px;background:#6366f1;pointer-events:none;">View Card</button>';
        html += '</div>';
      });
      html += '</div>';
      wrap.innerHTML = html;

      wrap.querySelectorAll(".history-item").forEach((el) => {
        el.addEventListener("click", () => {
          const hid = el.dataset.hid;
          closeHistoryModal();
          openHistoricalCardModal(hid);
        });
      });
    } catch (e) {
      wrap.innerHTML = '<div style="color:#e11d48;padding:16px;text-align:center;">Failed to load history from server.</div>';
    }
  }

  async function openHistoricalCardModal(historyId) {
    try {
      const res = await fetch("/api/history/" + historyId + "/", {
        headers: { "X-Requested-With": "XMLHttpRequest" }
      });
      const data = await res.json();
      if (!data.ok || !data.snapshot) {
        alert("Failed to load historical rate card snapshot.");
        return;
      }

      const sheet = data.snapshot;
      const modal = document.getElementById("fullview-modal");
      if (!modal) return;

      const titleEl = document.getElementById("fullview-title");
      if (titleEl) titleEl.textContent = (sheet.productLabel || "Treasury Rates") + " (Historical)";

      const subTitleEl = document.getElementById("fullview-subtitle");
      if (subTitleEl) subTitleEl.textContent = "Effective Date: " + (fmtDateDisp(toISODate(sheet.effectiveDate)) || sheet.effectiveDate);

      const badgeEl = document.getElementById("fullview-badge");
      if (badgeEl) {
        badgeEl.textContent = "HISTORICAL DB RECORD";
        badgeEl.style.background = "#6366f1";
      }

      const gridWrap = document.getElementById("fullview-grid-wrap");
      if (gridWrap) {
        gridWrap.innerHTML = gridHTML(sheet, false, "modal");
      }

      modal.hidden = false;
      document.body.style.overflow = "hidden";
      setTimeout(autoFitModalScale, 50);
    } catch (e) {
      alert("Error loading historical snapshot.");
    }
  }

  function closeHistoryModal() {
    const modal = document.getElementById("history-modal");
    if (modal) modal.hidden = true;
    document.body.style.overflow = "";
  }

  const btnViewFull = document.getElementById("btn-view-full");
  if (btnViewFull) btnViewFull.addEventListener("click", openFullViewModal);

  const btnNominal = document.getElementById("btn-nominal");
  if (btnNominal) btnNominal.addEventListener("click", toggleNominalMode);

  const btnFullDownload = document.getElementById("btn-fullview-download");
  if (btnFullDownload) btnFullDownload.addEventListener("click", downloadNominalMonthlyCSV);

  const btnFullClose = document.getElementById("btn-fullview-close");
  if (btnFullClose) btnFullClose.addEventListener("click", closeFullViewModal);

  const fullBackdrop = document.getElementById("fullview-backdrop");
  if (fullBackdrop) fullBackdrop.addEventListener("click", closeFullViewModal);

  const btnHistory = document.getElementById("btn-history");
  if (btnHistory) btnHistory.addEventListener("click", openHistoryModal);

  const btnHistoryClose = document.getElementById("btn-history-close");
  if (btnHistoryClose) btnHistoryClose.addEventListener("click", closeHistoryModal);

  const historyBackdrop = document.getElementById("history-backdrop");
  if (historyBackdrop) historyBackdrop.addEventListener("click", closeHistoryModal);

  const btnZoomIn = document.getElementById("btn-fullview-zoom-in");
  if (btnZoomIn) btnZoomIn.addEventListener("click", () => setModalScale(modalScale + 0.08));

  const btnZoomOut = document.getElementById("btn-fullview-zoom-out");
  if (btnZoomOut) btnZoomOut.addEventListener("click", () => setModalScale(modalScale - 0.08));

  const btnFit = document.getElementById("btn-fullview-fit");
  if (btnFit) btnFit.addEventListener("click", autoFitModalScale);

  window.addEventListener("resize", () => {
    const modal = document.getElementById("fullview-modal");
    if (modal && !modal.hidden) autoFitModalScale();
  });

  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      const modal = document.getElementById("fullview-modal");
      if (modal && !modal.hidden) closeFullViewModal();
      const hmodal = document.getElementById("history-modal");
      if (hmodal && !hmodal.hidden) closeHistoryModal();
    }
  });

  // setInterval(pollRates, 3000);

  render();
})();
