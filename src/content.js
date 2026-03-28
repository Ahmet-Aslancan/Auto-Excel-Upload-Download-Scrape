(function initContentScript() {
  // Prevent duplicate initialization when script is injected multiple times in same tab.
  if (window.__penExcelUpdaterContentScriptLoaded) {
    return;
  }
  window.__penExcelUpdaterContentScriptLoaded = true;

  // PEN ID and AWC Child ID are the same value; Excel may use either column name.
  const PEN_KEYS = [
    "PEN ID",
    "PEN_ID",
    "pen_id",
    "PEN",
    "pen",
    "AWC Child ID",
    "AWC_Child_ID",
    "awc_child_id",
    "AWC Child",
    "awcchildid"
  ];

  function normalizeKey(value) {
    return String(value || "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizePenId(value) {
    return String(value || "")
      .toUpperCase()
      .replace(/\s+/g, "")
      .trim();
  }

  function getLabelForElement(el) {
    if (!el) return "";
    const id = el.id;
    if (id) {
      const labelByFor = document.querySelector(`label[for="${CSS.escape(id)}"]`);
      if (labelByFor?.textContent) return normalizeKey(labelByFor.textContent);
    }

    const parentLabel = el.closest("label");
    if (parentLabel?.textContent) return normalizeKey(parentLabel.textContent);

    const prev = el.previousElementSibling;
    if (prev?.tagName === "LABEL" && prev.textContent) return normalizeKey(prev.textContent);

    return "";
  }

  function readElementValue(el) {
    const tag = el.tagName.toLowerCase();
    if (tag === "select") {
      return normalizeKey(el.options[el.selectedIndex]?.text || el.value || "");
    }
    if (el.type === "checkbox" || el.type === "radio") {
      return el.checked ? "true" : "";
    }
    return normalizeKey(el.value || el.getAttribute("value") || "");
  }

  function inferFieldName(el) {
    return (
      getLabelForElement(el) ||
      normalizeKey(el.getAttribute("aria-label")) ||
      normalizeKey(el.name) ||
      normalizeKey(el.id) ||
      normalizeKey(el.getAttribute("placeholder")) ||
      ""
    );
  }

  function extractPenId() {
    const penSelectors = [
      '[name="penId"]',
      '[name="pen_id"]',
      '[name="pen-id"]',
      "#penId",
      "#pen_id",
      '[data-pen-id]',
      '[id*="pen"][id*="id"]',
      '[name*="pen"][name*="id"]',
      '[name="awcChildId"]',
      '[name="awc_child_id"]',
      '#awcChildId',
      '#awc_child_id',
      '[id*="awc"][id*="child"][id*="id"]',
      '[name*="awc"][name*="child"][name*="id"]'
    ];

    for (const selector of penSelectors) {
      const el = document.querySelector(selector);
      if (!el) continue;

      const value =
        normalizeKey(el.value) ||
        normalizeKey(el.textContent) ||
        normalizeKey(el.getAttribute("data-pen-id"));
      if (value) return value;
    }

    const allLabels = Array.from(document.querySelectorAll("label, dt, th, strong, b, span, p"));
    for (const node of allLabels) {
      const text = normalizeKey(node.textContent);
      const isPenId = /pen\s*id/i.test(text);
      const isAwcChildId = /awc\s*child\s*id/i.test(text);
      if (!isPenId && !isAwcChildId) continue;

      const parentText = normalizeKey(node.parentElement?.textContent || "");
      const match =
        parentText.match(/pen\s*id\s*[:#-]?\s*([A-Za-z0-9_-]+)/i) ||
        parentText.match(/awc\s*child\s*id\s*[:#-]?\s*([A-Za-z0-9_-]+)/i);
      if (match?.[1]) return match[1];

      const sibling = node.nextElementSibling;
      if (sibling) {
        const siblingValue =
          normalizeKey(sibling.value) || normalizeKey(sibling.textContent) || normalizeKey(sibling.getAttribute("value"));
        if (siblingValue) return siblingValue;
      }
    }

    const bodyText = normalizeKey(document.body?.innerText || "");
    const bodyMatch =
      bodyText.match(/pen\s*id\s*[:#-]?\s*([A-Za-z0-9_-]+)/i) ||
      bodyText.match(/awc\s*child\s*id\s*[:#-]?\s*([A-Za-z0-9_-]+)/i);
    return bodyMatch?.[1] || "";
  }

  function scrapeFormFields() {
    const fields = {};
    const elements = Array.from(document.querySelectorAll("input, textarea, select"));
    for (const el of elements) {
      if (el.type === "hidden") continue;
      const key = inferFieldName(el);
      if (!key) continue;
      const value = readElementValue(el);
      if (!value) continue;
      fields[key] = value;
    }
    return fields;
  }

  function normalizeHeaderKey(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");
  }

  function readHeadersFromTable(table) {
    const theadHeaders = Array.from(table.querySelectorAll("thead th")).map((th) => normalizeKey(th.textContent));
    if (theadHeaders.length) return theadHeaders;

    const firstRow = table.querySelector("tr");
    if (!firstRow) return [];
    return Array.from(firstRow.querySelectorAll("th, td")).map((cell) => normalizeKey(cell.textContent));
  }

  function scoreStudentTable(headers) {
    const normalized = headers.map((h) => normalizeHeaderKey(h));
    let score = 0;
    if (normalized.includes("penid") || normalized.includes("awcchildid")) score += 5;
    if (normalized.includes("name")) score += 2;
    if (normalized.includes("sn")) score += 2;
    if (normalized.includes("gender")) score += 1;
    if (normalized.includes("dob")) score += 1;
    return score;
  }

  function extractRowsFromTable(table, headers) {
    const tbodyRows = Array.from(table.querySelectorAll("tbody tr"));
    const rawRows =
      tbodyRows.length > 0
        ? tbodyRows
        : Array.from(table.querySelectorAll("tr")).slice(1);

    const rows = [];
    for (const tr of rawRows) {
      if (tr.querySelectorAll("th").length) continue;
      const cells = Array.from(tr.querySelectorAll("td"));
      if (!cells.length) continue;

      const values = cells.map((td) => normalizeKey(td.textContent));
      if (!values.some((v) => v !== "")) continue;

      const rowObj = {};
      for (let i = 0; i < headers.length; i += 1) {
        const key = headers[i] || `Column ${i + 1}`;
        rowObj[key] = values[i] ?? "";
      }
      rows.push(rowObj);
    }
    return rows;
  }

  function extractStudentTableData() {
    const tables = Array.from(document.querySelectorAll("table"));
    if (!tables.length) {
      throw new Error("No table found on this page.");
    }

    let best = null;
    let bestScore = -1;

    for (let i = 0; i < tables.length; i += 1) {
      const table = tables[i];
      const headers = readHeadersFromTable(table);
      if (!headers.length) continue;
      const score = scoreStudentTable(headers);
      if (score > bestScore) {
        bestScore = score;
        best = { table, headers, index: i };
      }
    }

    if (!best || bestScore < 2) {
      throw new Error("Could not find a student table (needs columns like PEN ID / AWC Child ID / Name / SN).");
    }

    const headers = best.headers.map((h, i) => h || `Column ${i + 1}`);
    const rows = extractRowsFromTable(best.table, headers);
    if (!rows.length) {
      throw new Error("Student table found, but no data rows are visible.");
    }

    return {
      tableIndex: best.index,
      headers,
      rows
    };
  }

  async function extractStudentTableDataAllPages(options = {}) {
    const limitPages =
      Number.isInteger(options?.limitPages) && options.limitPages > 0 ? options.limitPages : null;

    const firstCtx = getCurrentTablePageContext();
    if (!firstCtx || !firstCtx.tableCtx) {
      // Fall back to the single-page extractor for a clearer error
      return extractStudentTableData();
    }

    await goToFirstTablePage();

    const collected = [];
    const seen = new Set();
    const headers = firstCtx.headers;
    const penColIndex = firstCtx.penColIndex;

    let pageCount = 0;
    while (true) {
      pageCount += 1;
      const ctx = getCurrentTablePageContext();
      if (!ctx || !ctx.tableCtx) break;

      const rows = ctx.rows || [];
      for (const row of rows) {
        const penIdRaw = penColIndex >= 0 ? row[headers[penColIndex]] : getPenIdFromRowData(row);
        const penId = normalizePenId(penIdRaw);
        const key = penId || JSON.stringify(row);
        if (seen.has(key)) continue;
        seen.add(key);
        collected.push(row);
      }

      if (limitPages && pageCount >= limitPages) break;

      const moved = await goToNextTablePageIfAvailable(seen);
      if (!moved.moved) break;
      await sleep(250);
    }

    if (!collected.length) {
      throw new Error("Student table found, but no data rows are visible.");
    }

    return {
      tableIndex: firstCtx.tableCtx.index,
      headers,
      rows: collected,
      pages: pageCount
    };
  }

  function getBestStudentTableContext() {
    const tables = Array.from(document.querySelectorAll("table"));
    if (!tables.length) return null;

    let best = null;
    let bestScore = -1;
    for (let i = 0; i < tables.length; i += 1) {
      const table = tables[i];
      const headers = readHeadersFromTable(table);
      if (!headers.length) continue;
      const score = scoreStudentTable(headers);
      if (score > bestScore) {
        bestScore = score;
        best = { table, headers, index: i };
      }
    }
    if (!best || bestScore < 2) return null;
    return best;
  }

  function getRowsForTable(table) {
    const tbodyRows = Array.from(table.querySelectorAll("tbody tr"));
    const allRows = tbodyRows.length ? tbodyRows : Array.from(table.querySelectorAll("tr")).slice(1);
    return allRows.filter((row) => isElementVisible(row));
  }

  function getCellsFromRow(rowEl) {
    return Array.from(rowEl.querySelectorAll("td"));
  }

  function getRowDataFromCells(headers, rowEl) {
    const cells = Array.from(rowEl.querySelectorAll("td"));
    const data = {};
    for (let i = 0; i < headers.length; i += 1) {
      const key = headers[i] || `Column ${i + 1}`;
      data[key] = normalizeKey(cells[i]?.textContent || "");
    }
    return data;
  }

  function findColumnIndex(headers, aliases) {
    const wanted = aliases.map((v) => normalizeHeaderKey(v));
    for (let i = 0; i < headers.length; i += 1) {
      const normalized = normalizeHeaderKey(headers[i]);
      if (wanted.includes(normalized)) return i;
    }
    return -1;
  }

  function findActionElementInRow(rowEl, screeningColIndex) {
    const cells = Array.from(rowEl.querySelectorAll("td"));
    if (screeningColIndex >= 0 && cells[screeningColIndex]) {
      const candidate = cells[screeningColIndex].querySelector(
        "a, button, input[type='button'], input[type='submit'], [role='button']"
      );
      if (candidate) return candidate;
    }

    const clickableInRow = Array.from(
      rowEl.querySelectorAll("a, button, input[type='button'], input[type='submit'], [role='button']")
    );
    const byText = clickableInRow.find((el) => /screening|action|details|view/i.test(normalizeKey(el.textContent || el.value)));
    if (byText) return byText;
    return clickableInRow[0] || null;
  }

  function findLiveRowByPenId(table, penId, penColIndex) {
    const target = normalizePenId(penId);
    if (!target) return null;
    const rows = getRowsForTable(table);

    for (const row of rows) {
      if (row.querySelectorAll("th").length) continue;
      const cells = getCellsFromRow(row);
      if (!cells.length) continue;

      const penText =
        penColIndex >= 0
          ? normalizePenId(cells[penColIndex]?.textContent || "")
          : normalizePenId(row.textContent || "");
      if (penText && penText === target) return row;
    }
    return null;
  }

  function extractVisibleRowsFromTable(table, headers) {
    const rawRows = getRowsForTable(table);
    const rows = [];

    for (const tr of rawRows) {
      if (tr.querySelectorAll("th").length) continue;
      const cells = Array.from(tr.querySelectorAll("td"));
      if (!cells.length) continue;

      const values = cells.map((td) => normalizeKey(td.textContent));
      if (!values.some((v) => v !== "")) continue;

      const rowObj = {};
      for (let i = 0; i < headers.length; i += 1) {
        const key = headers[i] || `Column ${i + 1}`;
        rowObj[key] = values[i] ?? "";
      }
      rows.push(rowObj);
    }

    return rows;
  }

  function getElementActionLabel(el) {
    const descendants = Array.from(el?.querySelectorAll?.("*") || []);
    const descendantHints = descendants.flatMap((node) => [
      node?.textContent,
      node?.getAttribute?.("title"),
      node?.getAttribute?.("aria-label"),
      node?.getAttribute?.("data-original-title"),
      node?.getAttribute?.("data-bs-original-title"),
      node?.getAttribute?.("ng-reflect-message"),
      node?.getAttribute?.("mattooltip"),
      node?.getAttribute?.("ng-reflect-mat-tooltip"),
      node?.getAttribute?.("tooltip")
    ]);

    return normalizeKey(
      [
        el?.textContent,
        el?.value,
        el?.getAttribute?.("title"),
        el?.getAttribute?.("aria-label"),
        el?.getAttribute?.("data-original-title"),
        el?.getAttribute?.("data-bs-original-title"),
        el?.getAttribute?.("ng-reflect-message"),
        el?.getAttribute?.("mattooltip"),
        el?.getAttribute?.("ng-reflect-mat-tooltip"),
        el?.getAttribute?.("tooltip"),
        ...descendantHints
      ]
        .filter(Boolean)
        .join(" ")
    );
  }

  function isElementVisible(el) {
    if (!(el instanceof Element)) return false;
    const style = window.getComputedStyle(el);
    if (style.display === "none" || style.visibility === "hidden" || Number(style.opacity) === 0) return false;
    const rect = el.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function getActionClassHints(el) {
    const selfClass = normalizeKey(el?.getAttribute?.("class") || "");
    const descendants = Array.from(el?.querySelectorAll?.("*") || []);
    const childClasses = descendants.map((node) => normalizeKey(node.getAttribute?.("class") || ""));
    return `${selfClass} ${childClasses.join(" ")}`.trim();
  }

  function getTooltipText(el) {
    const svgTitleText = Array.from(el?.querySelectorAll?.("svg title") || [])
      .map((n) => n.textContent)
      .filter(Boolean)
      .join(" ");

    return normalizeKey(
      [
        el?.getAttribute?.("title"),
        el?.getAttribute?.("aria-label"),
        el?.getAttribute?.("data-original-title"),
        el?.getAttribute?.("data-bs-original-title"),
        el?.getAttribute?.("ng-reflect-message"),
        el?.getAttribute?.("mattooltip"),
        el?.getAttribute?.("ng-reflect-mat-tooltip"),
        el?.getAttribute?.("tooltip"),
        svgTitleText
      ]
        .filter(Boolean)
        .join(" ")
    );
  }

  function getClickableAncestor(el) {
    if (!(el instanceof Element)) return null;
    return el.closest(
      "a, button, input[type='button'], input[type='submit'], [role='button'], [onclick], div[title], span[title], [class*='cursor-pointer']"
    );
  }

  function getScreeningStartDivFromScope(scopeEl) {
    if (!(scopeEl instanceof Element)) return null;
    const directExact = scopeEl.querySelector(
      "div[title='Screening Start'],div[title='screening start'],div[title='Start Screening'],div[title='start screening']"
    );
    if (directExact && isElementVisible(directExact)) return directExact;

    const direct = scopeEl.querySelector(
      "div[title*='Screening Start'],div[title*='Start Screening']"
    );
    if (direct && isElementVisible(direct)) return direct;

    const svgWithTitle = Array.from(scopeEl.querySelectorAll("svg")).find((svg) => {
      const title = normalizeKey(svg.querySelector("title")?.textContent || "");
      return /(screening\s*start|start\s*screening)/i.test(title) && !/edit\s*screening/i.test(title);
    });
    if (svgWithTitle) {
      const parentDiv = svgWithTitle.closest("div[title],div[class*='cursor-pointer'],button,a,[role='button']");
      if (parentDiv && isElementVisible(parentDiv)) return parentDiv;
    }
    return null;
  }

  function getScreeningActionsCell(rowEl, screeningColIndex) {
    const cells = Array.from(rowEl.querySelectorAll("td"));
    if (!cells.length) return null;
    if (screeningColIndex >= 0 && cells[screeningColIndex]) return cells[screeningColIndex];
    return cells[cells.length - 1] || null;
  }

  function getScreeningActionsContainer(actionCell) {
    if (!(actionCell instanceof Element)) return null;
    const directChildDiv = Array.from(actionCell.children).find((el) => el.tagName?.toLowerCase?.() === "div");
    if (directChildDiv) return directChildDiv;
    const fallback = actionCell.querySelector("div.flex.items-center.gap-3.justify-center, div");
    if (fallback) return fallback;
    return null;
  }

  function findScreeningStartElementInRow(rowEl, screeningColIndex) {
    const actionCell = getScreeningActionsCell(rowEl, screeningColIndex);
    const actionContainer = getScreeningActionsContainer(actionCell);
    const candidateScope = actionContainer || actionCell || rowEl;

    // Exact path requested by user:
    // tbody tr -> td:last-child -> div(container) -> div[title="Screening Start"]
    const directFromLastTd = rowEl.querySelector(
      "td:last-child > div > div[title='Screening Start'], td:last-child > div > div[title='screening start'], td:last-child > div > div[title='Start Screening'], td:last-child > div > div[title='start screening']"
    );
    if (directFromLastTd) return directFromLastTd;

    const directTitleMatch = Array.from(
      rowEl.querySelectorAll("td:last-child > div > div[title], td:last-child [title]")
    ).find((el) => {
      const title = normalizeKey(el.getAttribute("title") || "");
      return /(screening\s*start|start\s*screening)/i.test(title) && !/edit\s*screening/i.test(title);
    });
    if (directTitleMatch) return directTitleMatch;

    // Absolute priority: exact element pattern from provided HTML.
    const exactTitledDiv = getScreeningStartDivFromScope(candidateScope);
    if (exactTitledDiv) return exactTitledDiv;

    const exactSvgTitle = Array.from(candidateScope.querySelectorAll("svg title")).find((n) =>
      /(screening\s*start|start\s*screening)/i.test(normalizeKey(n.textContent)) &&
      !/edit\s*screening/i.test(normalizeKey(n.textContent))
    );
    if (exactSvgTitle) {
      const svg = exactSvgTitle.closest("svg");
      const exactParent =
        (svg && svg.closest("div[title],button,a,[role='button'],div[class*='cursor-pointer']")) ||
        getClickableAncestor(svg || exactSvgTitle);
      if (exactParent && isElementVisible(exactParent)) return exactParent;
    }

    // Highest priority: exact titled option (matches provided HTML).
    const titledNodes = Array.from(
      candidateScope.querySelectorAll(
        "[title],[aria-label],[data-original-title],[data-bs-original-title],[ng-reflect-message],[mattooltip],[ng-reflect-mat-tooltip],[tooltip]"
      )
    ).filter((el) => {
      if (!isElementVisible(el)) return false;
      const tip = getTooltipText(el);
      return /(screening\s*start|start\s*screening)/i.test(tip) && !/edit\s*screening/i.test(tip);
    });

    for (const node of titledNodes) {
      const clickableTarget = getClickableAncestor(node);
      if (clickableTarget && isElementVisible(clickableTarget)) return clickableTarget;
      if (isElementVisible(node)) return node;
    }

    const clickable = Array.from(
      candidateScope.querySelectorAll(
        "a, button, input[type='button'], input[type='submit'], [role='button'], [onclick], div[title], span[title], [class*='cursor-pointer']"
      )
    ).filter(isElementVisible);

    // Highest priority: exact SVG <title>Screening Start</title> match.
    const svgTitleNodes = Array.from(candidateScope.querySelectorAll("svg title")).filter((n) =>
      /(screening\s*start|start\s*screening)/i.test(normalizeKey(n.textContent)) &&
      !/edit\s*screening/i.test(normalizeKey(n.textContent))
    );
    for (const titleNode of svgTitleNodes) {
      const svg = titleNode.closest("svg");
      const clickableTarget = getClickableAncestor(svg || titleNode);
      if (clickableTarget && isElementVisible(clickableTarget)) return clickableTarget;
      if (svg && isElementVisible(svg)) return svg;
    }

    if (!clickable.length) return null;

    // Highest priority: click option carrying "Screening Start" hover text.
    const hoverCandidates = Array.from(
      candidateScope.querySelectorAll(
        "[title],[aria-label],[data-original-title],[data-bs-original-title],[ng-reflect-message],[mattooltip],[ng-reflect-mat-tooltip],[tooltip]"
      )
    ).filter((el) => {
      if (!isElementVisible(el)) return false;
      const tip = getTooltipText(el);
      return /(screening\s*start|start\s*screening)/i.test(tip) && !/edit\s*screening/i.test(tip);
    });

    for (const el of hoverCandidates) {
      const clickableTarget = getClickableAncestor(el);
      if (clickableTarget && isElementVisible(clickableTarget)) return clickableTarget;
      if (isElementVisible(el)) return el;
    }

    let bestEl = null;
    let bestScore = Number.NEGATIVE_INFINITY;

    clickable.forEach((el, idx) => {
      const label = getElementActionLabel(el);
      const classHints = getActionClassHints(el);
      let score = 0;

      if (/(screening\s*start|start\s*screening)/i.test(label)) score += 130;
      if (/start\s*screening/i.test(label)) score += 60;
      if (/screen/i.test(label)) score += 30;
      if (/start|play|launch|begin|fa-play|icon-start/i.test(classHints)) score += 25;
      if (/warning|yellow|orange|btn-warning/i.test(classHints)) score += 18;

      if (/edit\s*screening/i.test(label)) score -= 120;
      if (/pencil|edit|fa-edit|fa-pencil|bi-pencil|icon-edit/i.test(classHints)) score -= 80;
      if (/view|history|print|pdf|download|delete|remove/i.test(label)) score -= 70;
      if (/fa-eye|icon-eye|history|clock|pdf|print/i.test(classHints)) score -= 45;

      if (idx === 0) score += 8;

      if (score > bestScore) {
        bestScore = score;
        bestEl = el;
      }
    });

    if (bestEl && bestScore > 0) return bestEl;

    // Never click a generic fallback that may be "Edit Screening".
    return null;
  }

  function getLegacyScreeningDivFromScope(scopeEl) {
    if (!(scopeEl instanceof Element)) return null;
    const directExact = scopeEl.querySelector(
      "div[title='Edit Screening'],div[title='Edit screening'],div[title='edit screening'],div[title='Start Screening'],div[title='Start screening'],div[title='start screening']"
    );
    if (directExact && isElementVisible(directExact)) return directExact;

    const direct = scopeEl.querySelector(
      "div[title*='Edit Screening'],div[title*='Start Screening']"
    );
    if (direct && isElementVisible(direct)) return direct;

    const svgWithTitle = Array.from(scopeEl.querySelectorAll("svg")).find((svg) => {
      const title = normalizeKey(svg.querySelector("title")?.textContent || "");
      return /(edit|start)\s*screening/i.test(title);
    });
    if (svgWithTitle) {
      const parentDiv = svgWithTitle.closest("div[title],div[class*='cursor-pointer'],button,a,[role='button']");
      if (parentDiv && isElementVisible(parentDiv)) return parentDiv;
    }
    return null;
  }

  function findLegacyScreeningElementInRow(rowEl, screeningColIndex) {
    const actionCell = getScreeningActionsCell(rowEl, screeningColIndex);
    const actionContainer = getScreeningActionsContainer(actionCell);
    let candidateScope = actionContainer || actionCell || rowEl;

    const allCells = Array.from(rowEl.querySelectorAll("td"));
    for (const cell of allCells) {
      const scope = getScreeningActionsContainer(cell) || cell;
      const byTitle = scope.querySelector(
        "div[title*='Edit Screening'],div[title*='Start Screening'],div[title*='edit screening'],div[title*='start screening']"
      );
      if (byTitle && isElementVisible(byTitle)) return byTitle;
      const bySvgTitle = Array.from(scope.querySelectorAll("svg title")).find((n) =>
        /(edit|start)\s*screening/i.test(normalizeKey(n.textContent))
      );
      if (bySvgTitle) {
        const clickable = getClickableAncestor(bySvgTitle.closest("svg") || bySvgTitle);
        if (clickable && isElementVisible(clickable)) return clickable;
      }
      const byText = Array.from(scope.querySelectorAll("a, button, [role='button'], div, span")).find((el) => {
        if (!isElementVisible(el)) return false;
        const t = (el.textContent || "").replace(/\s+/g, " ").trim();
        return /^(edit\s*screening|start\s*screening)$/i.test(t) || t === "Edit Screening" || t === "Start Screening";
      });
      if (byText) return getClickableAncestor(byText) || byText;
    }

    const directFromLastTd = rowEl.querySelector(
      "td:last-child > div > div[title='Edit Screening'], td:last-child > div > div[title='Edit screening'], td:last-child > div > div[title='edit screening'], td:last-child > div > div[title='Start Screening'], td:last-child > div > div[title='Start screening'], td:last-child > div > div[title='start screening']"
    );
    if (directFromLastTd) return directFromLastTd;

    const directTitleMatch = Array.from(
      rowEl.querySelectorAll("td:last-child > div > div[title], td:last-child [title]")
    ).find((el) => /(edit|start)\s*screening/i.test(normalizeKey(el.getAttribute("title") || "")));
    if (directTitleMatch) return directTitleMatch;

    const exactTitledDiv = getLegacyScreeningDivFromScope(candidateScope);
    if (exactTitledDiv) return exactTitledDiv;

    const exactSvgTitle = Array.from(candidateScope.querySelectorAll("svg title")).find((n) =>
      /(edit|start)\s*screening/i.test(normalizeKey(n.textContent))
    );
    if (exactSvgTitle) {
      const svg = exactSvgTitle.closest("svg");
      const exactParent =
        (svg && svg.closest("div[title],button,a,[role='button'],div[class*='cursor-pointer']")) ||
        getClickableAncestor(svg || exactSvgTitle);
      if (exactParent && isElementVisible(exactParent)) return exactParent;
    }

    const titledNodes = Array.from(
      candidateScope.querySelectorAll(
        "[title],[aria-label],[data-original-title],[data-bs-original-title],[ng-reflect-message],[mattooltip],[ng-reflect-mat-tooltip],[tooltip]"
      )
    ).filter((el) => isElementVisible(el) && /edit\s*screening/i.test(getTooltipText(el)));

    for (const node of titledNodes) {
      const clickableTarget = getClickableAncestor(node);
      if (clickableTarget && isElementVisible(clickableTarget)) return clickableTarget;
      if (isElementVisible(node)) return node;
    }

    const clickable = Array.from(
      candidateScope.querySelectorAll(
        "a, button, input[type='button'], input[type='submit'], [role='button'], [onclick], div[title], span[title], [class*='cursor-pointer']"
      )
    ).filter(isElementVisible);

    const svgTitleNodes = Array.from(candidateScope.querySelectorAll("svg title")).filter((n) =>
      /(edit|start)\s*screening/i.test(normalizeKey(n.textContent))
    );
    for (const titleNode of svgTitleNodes) {
      const svg = titleNode.closest("svg");
      const clickableTarget = getClickableAncestor(svg || titleNode);
      if (clickableTarget && isElementVisible(clickableTarget)) return clickableTarget;
      if (svg && isElementVisible(svg)) return svg;
    }

    if (!clickable.length) return null;

    const hoverCandidates = Array.from(
      candidateScope.querySelectorAll(
        "[title],[aria-label],[data-original-title],[data-bs-original-title],[ng-reflect-message],[mattooltip],[ng-reflect-mat-tooltip],[tooltip]"
      )
    ).filter((el) => isElementVisible(el) && /(edit|start)\s*screening/i.test(getTooltipText(el)));

    for (const el of hoverCandidates) {
      const clickableTarget = getClickableAncestor(el);
      if (clickableTarget && isElementVisible(clickableTarget)) return clickableTarget;
      if (isElementVisible(el)) return el;
    }

    let bestEl = null;
    let bestScore = Number.NEGATIVE_INFINITY;

    clickable.forEach((el, idx) => {
      const label = getElementActionLabel(el);
      const classHints = getActionClassHints(el);
      let score = 0;

      if (/(edit|start)\s*screening/i.test(label)) score += 120;
      if (/start\s*screening/i.test(label)) score += 55;
      if (/\bedit\b/i.test(label)) score += 45;
      if (/screen/i.test(label)) score += 30;
      if (/pencil|edit|fa-edit|fa-pencil|bi-pencil|icon-edit/i.test(classHints)) score += 50;
      if (/warning|yellow|orange|btn-warning/i.test(classHints)) score += 18;

      if (/view|history|print|pdf|download|delete|remove/i.test(label)) score -= 70;
      if (/fa-eye|icon-eye|history|clock|pdf|print/i.test(classHints)) score -= 45;

      if (idx === 0) score += 8;

      if (score > bestScore) {
        bestScore = score;
        bestEl = el;
      }
    });

    if (bestEl && bestScore > 0) return bestEl;

    const byText = Array.from(
      candidateScope.querySelectorAll("a, button, [role='button'], div[class*='cursor-pointer'], span[class*='cursor-pointer'], div[title], span[title]")
    ).find((el) => {
      if (!isElementVisible(el)) return false;
      const text = (el.textContent || "").replace(/\s+/g, " ").trim();
      return /^(edit\s*screening|start\s*screening)$/i.test(text) ||
        /\b(edit\s*screening|start\s*screening)\b/i.test(text);
    });
    if (byText) return byText;

    const byTextDescendant = Array.from(candidateScope.querySelectorAll("a, button, [role='button'], div, span")).find((el) => {
      if (!isElementVisible(el)) return false;
      const directText = Array.from(el.childNodes)
        .filter((n) => n.nodeType === Node.TEXT_NODE)
        .map((n) => (n.textContent || "").replace(/\s+/g, " ").trim())
        .join(" ");
      if (/^(edit\s*screening|start\s*screening)$/i.test(directText)) return true;
      const full = (el.textContent || "").replace(/\s+/g, " ").trim();
      return full === "Edit Screening" || full === "Start Screening";
    });
    if (byTextDescendant) {
      const clickableTarget = getClickableAncestor(byTextDescendant);
      return clickableTarget && isElementVisible(clickableTarget) ? clickableTarget : byTextDescendant;
    }

    return clickable[0] || null;
  }

  function dispatchClickSequence(targetEl) {
    if (!(targetEl instanceof Element)) return;
    targetEl.scrollIntoView({ block: "center", behavior: "instant" });
    if (typeof targetEl.click === "function") {
      targetEl.click();
    }
    const events = ["pointerdown", "mousedown", "pointerup", "mouseup", "click"];
    for (const type of events) {
      targetEl.dispatchEvent(
        new MouseEvent(type, {
          bubbles: true,
          cancelable: true,
          composed: true,
          view: window
        })
      );
    }
  }

  function isDetailsPageReady() {
    return location.href !== "" && (hasScreeningDetailsSection() || hasDetailsFieldMarkers());
  }

  async function clickScreeningStartElement(actionEl) {
    // Strictly one click on the matching Screening Start target.
    clickElementOnce(actionEl);
    await sleep(140);
  }

  async function clickLegacyScreeningElement(actionEl) {
    const td = actionEl?.closest?.("td");
    const container = getScreeningActionsContainer(td);
    if (container) {
      dispatchClickSequence(container);
      await sleep(120);
    }
    dispatchClickSequence(actionEl);
    await sleep(140);
  }

  function hasScreeningDetailsSection() {
    const nodes = Array.from(document.querySelectorAll("h1, h2, h3, h4, h5, legend, label, span, div, p"));
    return nodes.some((node) => /screening\s*details/i.test(normalizeKey(node.textContent)));
  }

  function hasDetailsFieldMarkers() {
    const text = normalizeKey(document.body?.innerText || "");
    return (
      /weight\s*\(?.*kg\)?/i.test(text) ||
      /height\s*\/?\s*length/i.test(text) ||
      /blood\s*pressure/i.test(text) ||
      /vision\s*-\s*left\s*eye/i.test(text)
    );
  }

  function isChildScreeningPage(url = location.href) {
    return /\/RBSK\/childScreening(?:[/?#]|$)/i.test(String(url || ""));
  }

  function parseNumericValue(value) {
    const str = String(value || "").replace(/,/g, "");
    const match = str.match(/-?\d+(\.\d+)?/);
    if (!match) return NaN;
    return Number(match[0]);
  }

  function calculateBmi(weightKg, heightCm) {
    const weight = parseNumericValue(weightKg);
    const height = parseNumericValue(heightCm);
    if (!Number.isFinite(weight) || !Number.isFinite(height) || height <= 0) return "";
    const meters = height / 100;
    if (meters <= 0) return "";
    return (weight / (meters * meters)).toFixed(2);
  }

  function classifyBmi(bmiValue) {
    const bmi = parseNumericValue(bmiValue);
    if (!Number.isFinite(bmi) || bmi <= 0) return "";
    if (bmi < 18.5) return "Underweight";
    if (bmi < 25) return "Normal";
    if (bmi < 30) return "Overweight";
    return "Obese";
  }

  function scrapeLabeledValues() {
    const out = {};
    const labels = Array.from(document.querySelectorAll("label, dt, th, strong, b, span"));
    for (const label of labels) {
      const key = normalizeKey(label.textContent);
      if (!key || key.length > 70) continue;
      if (!/(weight|height|length|bmi|classification|blood pressure|vision|defect)/i.test(key)) continue;

      const sibling = label.nextElementSibling;
      const fromSibling = normalizeKey(sibling?.value || sibling?.textContent || "");
      if (fromSibling) {
        out[key] = fromSibling;
        continue;
      }

      const parentText = normalizeKey(label.parentElement?.textContent || "");
      const keyEscaped = key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const match = parentText.match(new RegExp(`${keyEscaped}\\s*[:\\-]?\\s*(.+)$`, "i"));
      if (match?.[1]) {
        out[key] = normalizeKey(match[1]);
      }
    }
    return out;
  }

  function pickField(fields, aliases, options = {}) {
    const wanted = aliases.map((v) => normalizeHeaderKey(v));
    const excluded = (options.excludeAliases || []).map((v) => normalizeHeaderKey(v));

    // Priority 1: exact key matches.
    for (const [key, value] of Object.entries(fields)) {
      const normalized = normalizeHeaderKey(key);
      if (!normalized) continue;
      if (excluded.some((ex) => ex && normalized.includes(ex))) continue;
      if (wanted.some((w) => w && normalized === w)) return value;
    }

    // Priority 2: partial key matches.
    for (const [key, value] of Object.entries(fields)) {
      const normalized = normalizeHeaderKey(key);
      if (!normalized) continue;
      if (excluded.some((ex) => ex && normalized.includes(ex))) continue;
      if (wanted.some((w) => w && normalized.includes(w))) return value;
    }
    return "";
  }

  function extractBpReading(text) {
    const raw = String(text || "");
    const match = raw.match(/(\d{2,3})\s*\/\s*(\d{2,3})/);
    if (!match) return "";
    return `${match[1]}/${match[2]}`;
  }

  function extractBloodPressureFromFields(fields) {
    const entries = Object.entries(fields || {});

    // Priority 1: explicit Blood Pressure key (but not classification-like keys).
    for (const [key, value] of entries) {
      const normalizedKey = normalizeHeaderKey(key);
      if (!/(bloodpressure|^bp$)/.test(normalizedKey)) continue;
      if (/(classification|result|class)/.test(normalizedKey)) continue;
      const reading = extractBpReading(value);
      if (reading) return reading;
    }

    // Priority 2: combine systolic + diastolic style fields if present.
    let systolic = "";
    let diastolic = "";
    for (const [key, value] of entries) {
      const normalizedKey = normalizeHeaderKey(key);
      if (!systolic && /(systolic|bpsystolic)/.test(normalizedKey)) {
        const n = String(value || "").match(/\d{2,3}/);
        if (n) systolic = n[0];
      }
      if (!diastolic && /(diastolic|bpdiastolic)/.test(normalizedKey)) {
        const n = String(value || "").match(/\d{2,3}/);
        if (n) diastolic = n[0];
      }
    }
    if (systolic && diastolic) return `${systolic}/${diastolic}`;

    // Priority 3: any BP-like value tied to BP-related keys.
    for (const [key, value] of entries) {
      const normalizedKey = normalizeHeaderKey(key);
      if (!/(bloodpressure|bp|systolic|diastolic)/.test(normalizedKey)) continue;
      if (/(classification|result|class)/.test(normalizedKey)) continue;
      const reading = extractBpReading(value);
      if (reading) return reading;
    }

    return "";
  }

  function extractBloodPressureFromPageText() {
    const text = String(document.body?.innerText || "");
    const nearLabel = text.match(/blood\s*pressure[\s:\-]*([0-9]{2,3}\s*\/\s*[0-9]{2,3})/i);
    if (nearLabel?.[1]) return extractBpReading(nearLabel[1]);
    return "";
  }

  function extractBloodPressureFromControls() {
    const bpSelectors = [
      "input[name='bloodPressureValue']",
      "input[name*='bloodPressure']",
      "input[id*='bloodPressure']",
      "input[placeholder*='120/80']"
    ];

    for (const selector of bpSelectors) {
      const controls = Array.from(document.querySelectorAll(selector));
      for (const control of controls) {
        const directValue = normalizeKey(control.value || control.getAttribute("value") || "");
        const reading = extractBpReading(directValue);
        if (reading) return reading;
      }
    }
    return "";
  }

  function extractBmiValue(text) {
    const raw = String(text || "").trim();
    if (!raw) return "";
    const decimalMatch = raw.match(/\b(\d{1,2}(?:\.\d{1,2})?)\b/);
    if (decimalMatch?.[1]) return decimalMatch[1];
    const parsed = parseNumericValue(raw);
    if (!Number.isFinite(parsed) || parsed <= 0) return "";
    return parsed.toFixed(2);
  }

  function extractBmiCalculatedFromControls() {
    const bmiSelectors = [
      "input[name='bmi']",
      "input[name='bmiCalculated']",
      "input[id='bmi']",
      "input[name*='bmi'][name*='calculated']",
      "input[id*='bmi'][id*='calculated']",
      "input[aria-label*='BMI'][aria-label*='calculated']",
      "input[placeholder*='BMI']"
    ];

    for (const selector of bmiSelectors) {
      const controls = Array.from(document.querySelectorAll(selector));
      for (const control of controls) {
        const value = normalizeKey(control.value || control.getAttribute("value") || "");
        const bmi = extractBmiValue(value);
        if (bmi) return bmi;
      }
    }
    return "";
  }

  function extractValueFromControlNearLabel(labelRegex, valueParser) {
    const labels = Array.from(document.querySelectorAll("label"));
    for (const label of labels) {
      const text = normalizeKey(label.textContent || "");
      if (!labelRegex.test(text)) continue;

      const directControl =
        label.nextElementSibling?.matches?.("input, textarea, select")
          ? label.nextElementSibling
          : null;
      const scopedControl =
        directControl ||
        label.parentElement?.querySelector?.("input, textarea, select") ||
        label.closest("div")?.querySelector?.("input, textarea, select");
      if (!scopedControl) continue;

      const raw = normalizeKey(scopedControl.value || scopedControl.getAttribute("value") || "");
      const parsed = valueParser(raw);
      if (parsed) return parsed;
    }
    return "";
  }

  function extractBmiCalculatedFromPageText() {
    const text = String(document.body?.innerText || "");
    const nearCalculated = text.match(/bmi\s*\(?.*calculated.*\)?[\s:\-]*([0-9]{1,2}(?:\.[0-9]{1,2})?)/i);
    if (nearCalculated?.[1]) return nearCalculated[1];
    const nearBmi = text.match(/\bbmi\b[\s:\-]*([0-9]{1,2}(?:\.[0-9]{1,2})?)/i);
    if (nearBmi?.[1]) return nearBmi[1];
    return "";
  }

  function scrapeScreeningDetailsData() {
    const formFields = scrapeFormFields();
    const labeledValues = scrapeLabeledValues();
    const allFields = { ...labeledValues, ...formFields };

    const weight = pickField(allFields, ["weightkg", "weight"]);
    const height = pickField(allFields, ["heightlengthcm", "heightcm", "height", "lengthcm", "length"]);

    const bmiCalculatedFromPage =
      extractBmiCalculatedFromControls() ||
      extractValueFromControlNearLabel(/bmi\s*\(?.*calculated.*\)?/i, extractBmiValue) ||
      pickField(allFields, ["bmicalculated", "bmi"], {
        excludeAliases: ["classification", "result", "class"]
      }) ||
      extractBmiCalculatedFromPageText();
    const bmiResult = pickField(allFields, ["bmiresult"]);
    const bmiClassification = pickField(allFields, ["bmiclassification", "bmiclass"]);
    const bloodPressure =
      extractBloodPressureFromControls() ||
      extractBloodPressureFromFields(allFields) ||
      extractBloodPressureFromPageText() ||
      pickField(allFields, ["bloodpressure"], {
        excludeAliases: ["classification", "result", "class"]
      });
    const bloodPressureClassification = pickField(allFields, ["bloodpressureclassification"]);
    const leftEye = pickField(allFields, ["visionlefteye", "lefteyevision", "leftvision"]);
    const rightEye = pickField(allFields, ["visionrighteye", "righteyevision", "rightvision"]);
    const hbCount = pickField(allFields, ["hbcount", "hb"]);

    const bmiCalculatedFinal = bmiCalculatedFromPage || calculateBmi(weight, height);
    const bmiClassFinal = bmiClassification || classifyBmi(bmiCalculatedFinal);

    return {
      "Weight (kg)*": weight || "",
      "Height/Length (cm)*": height || "",
      "BMI (calculated)": bmiCalculatedFinal || "",
      "BMI Classification": bmiClassFinal || "",
      "BMI Result": bmiResult || "",
      "Blood Pressure*": bloodPressure || "",
      "Blood Pressure Classification": bloodPressureClassification || "",
      "Vision - Left Eye*": leftEye || "",
      "Vision - Right Eye *": rightEye || "",
      "Hb Count": hbCount || ""
    };
  }

  async function waitForStudentTableVisible(timeoutMs = 12000) {
    return waitFor(() => Boolean(getBestStudentTableContext()), timeoutMs, 120);
  }

  function findBackButton() {
    const clickable = Array.from(document.querySelectorAll("a, button, input[type='button'], [role='button']"));
    const candidate = clickable.find((el) => /back|return|list|previous/i.test(normalizeKey(el.textContent || el.value)));
    return candidate || null;
  }

  function findDetailsBackButton() {
    const clickable = Array.from(document.querySelectorAll("a, button, input[type='button'], [role='button']"));
    const byExactLabel = clickable.find((el) => /^back$/i.test(normalizeKey(el.textContent || el.value)));
    if (byExactLabel) return byExactLabel;

    const byStyleAndLabel = clickable.find((el) => {
      const label = normalizeKey(el.textContent || el.value);
      if (!/back/i.test(label)) return false;
      const className = normalizeKey(el.getAttribute("class") || "");
      return /bg-blue-50|text-blue-700|border-blue-200/.test(className);
    });
    if (byStyleAndLabel) return byStyleAndLabel;

    return findBackButton();
  }

  async function clickDetailsBackButtonAndPause(pauseMs = 1200) {
    const backButton = findDetailsBackButton();
    if (!backButton) {
      throw new Error('Back button not found on details page.');
    }

    dispatchClickSequence(backButton);
    await waitFor(
      () => isChildScreeningPage() && Boolean(getBestStudentTableContext()),
      10000,
      120
    );
    await sleep(pauseMs);
  }

  async function returnToStudentTable(baseChildScreeningUrl = "") {
    if (isChildScreeningPage() && (await waitForStudentTableVisible(1200))) {
      return;
    }

    const backButton = findDetailsBackButton();
    if (backButton) {
      dispatchClickSequence(backButton);
      const ok = await waitFor(
        () => isChildScreeningPage() && Boolean(getBestStudentTableContext()),
        10000,
        120
      );
      if (ok) return;
    }

    if (!isChildScreeningPage()) {
      history.back();
      const ok = await waitFor(
        () => isChildScreeningPage() && Boolean(getBestStudentTableContext()),
        10000,
        120
      );
      if (ok) return;
    }

    if (baseChildScreeningUrl && location.href !== baseChildScreeningUrl && !isChildScreeningPage()) {
      location.href = baseChildScreeningUrl;
      const ok = await waitFor(
        () => isChildScreeningPage() && Boolean(getBestStudentTableContext()),
        10000,
        120
      );
      if (ok) return;
    }

    throw new Error("Failed to return to childScreening table page.");
  }

  async function openScreeningDetailsForRow(rowEl, actionEl, options = {}) {
    const useLegacyClick = options?.useLegacyClick === true;
    const prevUrl = location.href;
    const anchor = actionEl.tagName.toLowerCase() === "a" ? actionEl : actionEl.closest("a");
    if (anchor && anchor.target === "_blank") {
      anchor.target = "_self";
    }

    rowEl.scrollIntoView({ block: "center", behavior: "instant" });
    if (useLegacyClick) {
      await clickLegacyScreeningElement(actionEl);
    } else {
      await clickScreeningStartElement(actionEl);
    }

    let opened = await waitFor(
      () => location.href !== prevUrl || isDetailsPageReady(),
      12000,
      120
    );

    // Fallback only if first click did not open details page.
    if (!opened) {
      const svg = actionEl?.querySelector?.("svg");
      if (svg) {
        dispatchClickSequence(svg);
        await sleep(200);
        opened = await waitFor(
          () => location.href !== prevUrl || isDetailsPageReady(),
          8000,
          120
        );
      }
    }

    if (!opened) throw new Error("Could not open Screening Details page from action button.");
    if (/\/RBSK\/reports\/remainingScreeningChildren(?:[/?#]|$)/i.test(location.href)) {
      throw new Error("Wrong navigation target detected (remainingScreeningChildren) instead of student details.");
    }
    // Hold briefly on details page so dynamic fields can render before any next step.
    await sleep(800);
  }

  function assertDetailsPenIdMatches(expectedPenId) {
    const expected = normalizePenId(expectedPenId);
    if (!expected) throw new Error("Missing PEN ID / AWC Child ID from Excel row.");
    const pagePen = normalizePenId(extractPenId());
    if (!pagePen) {
      throw new Error("Could not read PEN ID / AWC Child ID on Screening Details page for verification.");
    }
    if (pagePen !== expected) {
      throw new Error(`PEN ID / AWC Child ID mismatch. Excel "${expectedPenId}" does not match page "${pagePen}".`);
    }
  }

  function getCurrentTablePageContext(limitRows = null) {
    const tableCtx = getBestStudentTableContext();
    if (!tableCtx) return null;

    const headers = tableCtx.headers.map((h, i) => h || `Column ${i + 1}`);
    const actionIndex = findColumnIndex(headers, ["screening actions", "screening", "actions"]);
    const penColIndex = findColumnIndex(headers, ["pen id", "pen", "pen_id", "awc child id", "awcchildid"]);
    const rowsAll = extractVisibleRowsFromTable(tableCtx.table, headers);
    const rows = limitRows ? rowsAll.slice(0, limitRows) : rowsAll;

    return {
      tableCtx,
      headers,
      actionIndex,
      penColIndex,
      rows
    };
  }

  function getCurrentPagePenIds() {
    const ctx = getCurrentTablePageContext();
    if (!ctx) return [];
    return ctx.rows
      .map((row) => getPenIdFromRowData(row))
      .map((penId) => normalizePenId(penId))
      .filter(Boolean);
  }

  function getCurrentPageSignature() {
    return getCurrentPagePenIds().join("|");
  }

  function isDisabledButton(el) {
    if (!(el instanceof Element)) return true;
    if ("disabled" in el && el.disabled) return true;
    if (el.hasAttribute("disabled")) return true;
    const ariaDisabled = normalizeKey(el.getAttribute("aria-disabled") || "");
    if (ariaDisabled === "true") return true;
    return false;
  }

  function findNextPageButton() {
    const clickable = Array.from(document.querySelectorAll("button, a, input[type='button'], [role='button']")).filter(
      isElementVisible
    );
    if (!clickable.length) return null;

    const exact = clickable.find((el) => /^next$/i.test(normalizeKey(el.textContent || el.value)));
    if (exact) return exact;

    const byStyle = clickable.find((el) => {
      const label = normalizeKey(el.textContent || el.value);
      if (!/next/i.test(label)) return false;
      const className = normalizeKey(el.getAttribute("class") || "");
      return /bg-teal-500|border-teal-500|hover:bg-teal-700/.test(className);
    });
    if (byStyle) return byStyle;

    return clickable.find((el) => /\bnext\b/i.test(normalizeKey(el.textContent || el.value))) || null;
  }

  function findPreviousPageButton() {
    const clickable = Array.from(document.querySelectorAll("button, a, input[type='button'], [role='button']")).filter(
      isElementVisible
    );
    if (!clickable.length) return null;

    const exact = clickable.find((el) => /^previous$/i.test(normalizeKey(el.textContent || el.value)));
    if (exact) return exact;

    return clickable.find((el) => /\bprevious\b/i.test(normalizeKey(el.textContent || el.value))) || null;
  }

  async function goToFirstTablePage() {
    const prevButton = findPreviousPageButton();
    if (!prevButton) {
      const pager = getPaginationStateNearNext();
      if (pager && pager.current <= 1) return;
      return;
    }

    while (true) {
      const beforePager = getPaginationStateNearNext();
      if (beforePager && beforePager.current <= 1) return;
      if (isDisabledButton(prevButton)) return;

      const beforeSignature = getCurrentPageSignature();
      clickElementOnce(prevButton);

      const changed = await waitFor(() => {
        const afterPager = getPaginationStateNearNext();
        if (beforePager && afterPager && afterPager.current === 1) return true;
        if (afterPager && afterPager.current < (beforePager?.current ?? 999)) return true;
        const currentSignature = getCurrentPageSignature();
        if (Boolean(currentSignature) && currentSignature !== beforeSignature) return true;
        return false;
      }, 15000, 120);

      if (!changed) return;

      const afterPager = getPaginationStateNearNext();
      if (afterPager && afterPager.current <= 1) return;
    }
  }

  function clickElementOnce(targetEl) {
    if (!(targetEl instanceof Element)) return;
    targetEl.scrollIntoView({ block: "center", behavior: "instant" });
    if (typeof targetEl.click === "function") {
      targetEl.click();
      return;
    }
    targetEl.dispatchEvent(
      new MouseEvent("click", {
        bubbles: true,
        cancelable: true,
        composed: true,
        view: window
      })
    );
  }

  function findButtonLikeByText(pattern, scope = document) {
    const candidates = Array.from(
      scope.querySelectorAll("button, a, input[type='button'], input[type='submit'], [role='button']")
    ).filter((el) => isElementVisible(el));
    return (
      candidates.find((el) => {
        const text = normalizeKey(el.textContent || el.value || "");
        return pattern.test(text);
      }) || null
    );
  }

  function findPreviewSubmitButton() {
    return findButtonLikeByText(/^preview\s*&\s*submit$/i) || findButtonLikeByText(/preview\s*&\s*submit/i);
  }

  function findConfirmSubmitButton() {
    const modalScopes = Array.from(document.querySelectorAll("div.fixed.inset-0, [role='dialog'], .modal, .dialog"));
    for (const scope of modalScopes) {
      const btn = findButtonLikeByText(/confirm\s*&\s*submit/i, scope);
      if (btn) return btn;
    }
    return findButtonLikeByText(/confirm\s*&\s*submit/i);
  }

  async function submitDetailsThroughPreviewModal() {
    const previewBtn = findPreviewSubmitButton();
    if (!previewBtn) throw new Error('Preview & Submit button not found on details page.');

    clickElementOnce(previewBtn);

    const confirmVisible = await waitFor(() => Boolean(findConfirmSubmitButton()), 10000, 120);
    if (!confirmVisible) throw new Error('Confirm & Submit button did not appear in preview modal.');

    const confirmBtn = findConfirmSubmitButton();
    if (!confirmBtn) throw new Error('Confirm & Submit button not found in preview modal.');
    clickElementOnce(confirmBtn);

    await waitFor(() => !findConfirmSubmitButton(), 12000, 120);
    await sleep(500);
  }

  function parsePaginationFromText(text) {
    const raw = String(text || "");
    const slashMatch = raw.match(/\b(\d{1,4})\s*\/\s*(\d{1,4})\b/i);
    const ofMatch = raw.match(/\b(?:page\s*)?(\d{1,4})\s*(?:of)\s*(\d{1,4})\b/i);
    const match = slashMatch || ofMatch;
    if (!match) return null;
    const current = Number(match[1]);
    const total = Number(match[2]);
    if (!Number.isFinite(current) || !Number.isFinite(total) || total <= 0) return null;
    return { current, total };
  }

  function getPaginationStateNearNext() {
    const nextButton = findNextPageButton();
    if (!nextButton) return null;

    const scope =
      nextButton.closest("div.flex.flex-col") ||
      nextButton.closest("section") ||
      nextButton.parentElement ||
      document.body;
    if (!scope) return null;

    const candidates = [
      normalizeKey(scope.textContent || ""),
      ...Array.from(scope.querySelectorAll("div, span, p, small, strong, label")).map((el) => normalizeKey(el.textContent))
    ].filter(Boolean);

    for (const text of candidates) {
      const parsed = parsePaginationFromText(text);
      if (parsed) return parsed;
    }

    const wholePageParsed = parsePaginationFromText(normalizeKey(document.body?.innerText || ""));
    if (wholePageParsed) return wholePageParsed;
    return null;
  }

  async function goToNextTablePageIfAvailable(_processedPenIdsSet) {
    const nextButton = findNextPageButton();
    if (!nextButton || isDisabledButton(nextButton)) {
      return { moved: false, reason: "disabled_or_missing" };
    }

    const beforePager = getPaginationStateNearNext();
    if (beforePager && beforePager.current >= beforePager.total) {
      return { moved: false, reason: "last_page_reached" };
    }

    const beforeSignature = getCurrentPageSignature();
    clickElementOnce(nextButton);

    const changed = await waitFor(() => {
      const afterPager = getPaginationStateNearNext();
      if (beforePager && afterPager && afterPager.current === beforePager.current + 1) return true;
      const currentSignature = getCurrentPageSignature();
      if (Boolean(currentSignature) && currentSignature !== beforeSignature) return true;
      if (beforePager && afterPager && afterPager.current !== beforePager.current) return true;
      return false;
    }, 20000, 120);
    if (!changed) {
      return { moved: false, reason: "page_did_not_change" };
    }

    const afterPager = getPaginationStateNearNext();
    if (afterPager && afterPager.current > afterPager.total) {
      return { moved: false, reason: "invalid_pager_state" };
    }

    const rowsLoaded = await waitFor(() => getCurrentPagePenIds().length > 0, 10000, 120);
    if (!rowsLoaded) {
      return { moved: false, reason: "rows_not_loaded_after_next" };
    }

    const afterPenIds = getCurrentPagePenIds();
    if (!afterPenIds.length) {
      return { moved: false, reason: "no_rows_after_next" };
    }

    const afterSignature = afterPenIds.join("|");
    if (afterSignature === beforeSignature && !(beforePager && afterPager && afterPager.current > beforePager.current)) {
      return { moved: false, reason: "no_new_pen_ids" };
    }

    return { moved: true, reason: "ok" };
  }

  function getLiveRowMatchByPenId(penId) {
    const liveCtx = getBestStudentTableContext();
    if (!liveCtx) return null;
    const headers = liveCtx.headers.map((h, idx) => h || `Column ${idx + 1}`);
    const actionIndex = findColumnIndex(headers, ["screening actions", "screening", "actions"]);
    const penColIndex = findColumnIndex(headers, ["pen id", "pen", "pen_id", "awc child id", "awcchildid"]);
    const liveRow = findLiveRowByPenId(liveCtx.table, penId, penColIndex);
    if (!liveRow) return null;
    return {
      liveCtx,
      headers,
      actionIndex,
      penColIndex,
      liveRow
    };
  }

  async function findRowByPenIdAcrossPages(penId) {
    const normalizedTarget = normalizePenId(penId);
    if (!normalizedTarget) return null;

    // Check current page first.
    const currentMatch = getLiveRowMatchByPenId(normalizedTarget);
    if (currentMatch) return currentMatch;

    // If not found, keep moving to next page and check again.
    while (true) {
      const nextStep = await goToNextTablePageIfAvailable(new Set());
      if (!nextStep.moved) {
        await goToFirstTablePage();
        // Check page 1 after resetting (row may have been on first page if we started from a later page)
        const firstPageMatch = getLiveRowMatchByPenId(normalizedTarget);
        return firstPageMatch || null;
      }
      const movedMatch = getLiveRowMatchByPenId(normalizedTarget);
      if (movedMatch) return movedMatch;
    }
  }

  async function extractStudentTableWithDetails(options = {}) {
    const keepOnDetails = options?.keepOnDetails !== false;
    const limitRows = Number.isInteger(options?.limitRows) && options.limitRows > 0 ? options.limitRows : null;
    if (!isChildScreeningPage()) {
      throw new Error("Open the child screening table page first: /RBSK/childScreening");
    }
    const baseChildScreeningUrl = location.href;
    const outRows = [];
    const failures = [];
    const processedPenIds = new Set();
    const allHeaders = new Set();
    const detailHeaders = [
      "Weight (kg)*",
      "Height/Length (cm)*",
      "BMI (calculated)",
      "BMI Classification",
      "BMI Result",
      "Blood Pressure*",
      "Blood Pressure Classification",
      "Vision - Left Eye*",
      "Vision - Right Eye *",
      "Hb Count"
    ];

    while (true) {
      let pageCtx = getCurrentTablePageContext(limitRows);
      if (!pageCtx || !pageCtx.rows.length) {
        const ready = await waitFor(() => {
          const ctx = getCurrentTablePageContext(limitRows);
          return Boolean(ctx && ctx.rows.length);
        }, 10000, 120);
        if (!ready) break;
        pageCtx = getCurrentTablePageContext(limitRows);
      }
      if (!pageCtx) {
        throw new Error("Student table not found.");
      }
      for (const h of pageCtx.headers) allHeaders.add(h);

      const pendingRows = pageCtx.rows.filter((row) => {
        const penId = normalizePenId(getPenIdFromRowData(row));
        return Boolean(penId) && !processedPenIds.has(penId);
      });

      if (!pendingRows.length) {
        if (keepOnDetails) break;
        const nextStep = await goToNextTablePageIfAvailable(processedPenIds);
        if (!nextStep.moved) break;
        continue;
      }

      const baseRow = pendingRows[0];
      const penId = getPenIdFromRowData(baseRow);
      const normalizedPenId = normalizePenId(penId);

      try {
        const liveCtx = getBestStudentTableContext();
        if (!liveCtx) throw new Error("Student table not visible while processing rows.");
        const liveHeaders = liveCtx.headers.map((h, idx) => h || `Column ${idx + 1}`);
        const liveActionIndex = findColumnIndex(liveHeaders, ["screening actions", "screening", "actions"]);
        const livePenColIndex = findColumnIndex(liveHeaders, ["pen id", "pen", "pen_id", "awc child id", "awcchildid"]);
        const liveRow = findLiveRowByPenId(liveCtx.table, penId, livePenColIndex);
        if (!liveRow) throw new Error(`Row with PEN ID / AWC Child ID "${penId}" not found in current table.`);

        const actionEl = findLegacyScreeningElementInRow(liveRow, liveActionIndex);
        if (!actionEl) {
          throw new Error('Screening action button not found in this row.');
        }

        await openScreeningDetailsForRow(liveRow, actionEl, { useLegacyClick: true });
        const details = scrapeScreeningDetailsData();
        outRows.push({
          ...baseRow,
          ...details
        });
        processedPenIds.add(normalizedPenId);

        if (keepOnDetails) {
          for (const h of detailHeaders) allHeaders.add(h);
          await clickDetailsBackButtonAndPause(1200);
          return {
            headers: Array.from(allHeaders),
            rows: outRows,
            failures,
            extracted: outRows.length,
            pausedAfterBack: true,
            currentPenId: penId
          };
        }

        await returnToStudentTable(baseChildScreeningUrl);
      } catch (error) {
        processedPenIds.add(normalizedPenId);
        const reason = (error instanceof Error ? error.message : String(error)) || "Unknown extraction error.";
        failures.push({
          penId,
          reason: reason.trim() || "Unknown extraction error."
        });

        if (!keepOnDetails) {
          try {
            await returnToStudentTable(baseChildScreeningUrl);
          } catch (_returnErr) {
            // Continue with next row; page may already be table view.
          }
        }
      }
    }

    if (!outRows.length) {
      const firstFailure = (failures[0]?.reason || "").trim() || "Screening action (Edit Screening / Start Screening) not found or details page did not open.";
      throw new Error(`No student details could be extracted from Screening Actions. First issue: ${firstFailure}`);
    }

    for (const h of detailHeaders) allHeaders.add(h);

    return {
      headers: Array.from(allHeaders),
      rows: outRows,
      failures,
      extracted: outRows.length
    };
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  async function waitFor(checkFn, timeoutMs = 4000, intervalMs = 80) {
    const startedAt = Date.now();
    while (Date.now() - startedAt < timeoutMs) {
      if (checkFn()) return true;
      await sleep(intervalMs);
    }
    return false;
  }

  function getPenIdFromRowData(rowData) {
    for (const key of PEN_KEYS) {
      const value = rowData?.[key];
      if (String(value || "").trim()) return String(value).trim();
    }
    return "";
  }

  function getFieldHints(el) {
    if (!(el instanceof Element)) return "";
    const id = normalizeKey(el.getAttribute("id") || "");
    const name = normalizeKey(el.getAttribute("name") || "");
    const placeholder = normalizeKey(el.getAttribute("placeholder") || "");
    const ariaLabel = normalizeKey(el.getAttribute("aria-label") || "");
    const title = normalizeKey(el.getAttribute("title") || "");
    const label = getLabelForElement(el);
    return normalizeKey([label, name, id, placeholder, ariaLabel, title].filter(Boolean).join(" "));
  }

  function getRowValueByAliases(rowData, aliases) {
    const wanted = aliases.map((v) => normalizeHeaderKey(v));
    for (const [rawKey, rawValue] of Object.entries(rowData || {})) {
      const key = normalizeHeaderKey(rawKey);
      if (!key) continue;
      const val = String(rawValue ?? "").trim();
      if (!val) continue;
      if (wanted.some((alias) => key === alias || key.includes(alias))) {
        return val;
      }
    }
    return "";
  }

  function findFormControlByAliases(aliases) {
    const controls = Array.from(document.querySelectorAll("input, textarea, select")).filter((el) => {
      if (!(el instanceof Element)) return false;
      const inputType = normalizeKey(el.getAttribute("type") || "");
      if (inputType === "hidden") return false;
      if (el.hasAttribute("disabled")) return false;
      return isElementVisible(el);
    });
    if (!controls.length) return null;

    let best = null;
    let bestScore = Number.NEGATIVE_INFINITY;
    const wanted = aliases.map((v) => normalizeHeaderKey(v));

    for (const control of controls) {
      const hints = getFieldHints(control);
      const hintsNormalized = normalizeHeaderKey(hints);
      if (!hintsNormalized) continue;
      let score = 0;
      for (const alias of wanted) {
        if (!alias) continue;
        if (hintsNormalized === alias) score += 100;
        else if (hintsNormalized.includes(alias)) score += 55;
        else if (alias.includes(hintsNormalized)) score += 30;
      }
      if (score > bestScore) {
        bestScore = score;
        best = control;
      }
    }

    return bestScore > 0 ? best : null;
  }

  function setControlValue(control, value) {
    const text = String(value ?? "");
    const tag = control.tagName.toLowerCase();

    if (tag === "select") {
      const target = normalizeHeaderKey(text);
      const options = Array.from(control.options || []);
      const direct = options.find(
        (opt) =>
          normalizeHeaderKey(opt.textContent || "") === target || normalizeHeaderKey(opt.value || "") === target
      );
      const partial =
        direct ||
        options.find(
          (opt) =>
            normalizeHeaderKey(opt.textContent || "").includes(target) ||
            target.includes(normalizeHeaderKey(opt.textContent || "")) ||
            normalizeHeaderKey(opt.value || "").includes(target)
        );
      if (partial) {
        control.value = partial.value;
      } else {
        control.value = text;
      }
      control.dispatchEvent(new Event("input", { bubbles: true }));
      control.dispatchEvent(new Event("change", { bubbles: true }));
      return;
    }

    if (tag === "textarea") {
      const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value")?.set;
      if (nativeSetter) nativeSetter.call(control, text);
      else control.value = text;
      control.dispatchEvent(new Event("input", { bubbles: true }));
      control.dispatchEvent(new Event("change", { bubbles: true }));
      return;
    }

    const inputType = normalizeKey(control.getAttribute("type") || "");
    if (inputType === "checkbox" || inputType === "radio") {
      const truthy = /^(true|yes|1|checked)$/i.test(text);
      control.checked = truthy;
      control.dispatchEvent(new Event("input", { bubbles: true }));
      control.dispatchEvent(new Event("change", { bubbles: true }));
      return;
    }

    const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value")?.set;
    if (nativeSetter) nativeSetter.call(control, text);
    else control.value = text;
    control.dispatchEvent(new Event("input", { bubbles: true }));
    control.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function fillScreeningDetailsFormFromRow(rowData) {
    const fieldMap = [
      {
        excelAliases: ["weightkg", "weight(kg)", "weight"],
        controlAliases: ["weightkg", "weight"]
      },
      {
        excelAliases: ["heightlengthcm", "heightcm", "lengthcm", "height/length(cm)"],
        controlAliases: ["heightlengthcm", "heightcm", "height", "lengthcm", "length"]
      },
      {
        excelAliases: ["bmicalculated", "bmi(calculated)"],
        controlAliases: ["bmicalculated", "bmi"]
      },
      {
        excelAliases: ["bmiclassification"],
        controlAliases: ["bmiclassification", "bmi classification", "bmi category"]
      },
      {
        excelAliases: ["bmiresult", "bmiresul", "bmi result"],
        controlAliases: ["bmiresult", "bmi result"]
      },
      {
        excelAliases: ["bloodpressure", "bloodpressure*"],
        controlAliases: ["bloodpressure", "bp"]
      },
      {
        excelAliases: ["bloodpressureclassification", "blood pressure category"],
        controlAliases: ["bloodpressureclassification", "bpclassification", "blood pressure category"]
      },
      {
        excelAliases: ["visionlefteye", "vision-left-eye", "lefteyevision"],
        controlAliases: ["visionlefteye", "left eye", "lefteyevision"]
      },
      {
        excelAliases: ["visionrighteye", "vision-right-eye", "righteyevision"],
        controlAliases: ["visionrighteye", "right eye", "righteyevision"]
      },
      {
        excelAliases: ["hbcount", "hb", "hemoglobinlevel", "hemoglobin level"],
        controlAliases: ["hbcount", "hb", "hemoglobin", "hemoglobin level"]
      },
      {
        excelAliases: ["birthdefectfound", "birth defect found", "birthdefect"],
        controlAliases: ["birthdefectfound", "birth defect found", "birth defect"]
      }
    ];

    const missingFields = [];
    let filledCount = 0;

    for (const item of fieldMap) {
      const value = getRowValueByAliases(rowData, item.excelAliases);
      if (!String(value || "").trim()) continue;
      const control = findFormControlByAliases(item.controlAliases);
      if (!control) {
        missingFields.push(item.excelAliases[0]);
        continue;
      }
      setControlValue(control, value);
      filledCount += 1;
    }

    return {
      filledCount,
      missingFields
    };
  }

  async function settleMandatoryScreeningDetails(rowData, timeoutMs = 12000) {
    const weightCtrl = findFormControlByAliases(["weightkg", "weight"]);
    const heightCtrl = findFormControlByAliases(["heightlengthcm", "heightcm", "height", "lengthcm", "length"]);
    const bmiCtrl = findFormControlByAliases(["bmicalculated", "bmi", "bmi calculated"]);
    const bpCtrl = findFormControlByAliases(["bloodpressure", "bp", "blood pressure"]);
    const leftEyeCtrl = findFormControlByAliases(["visionlefteye", "left eye", "lefteyevision", "vision - left eye"]);
    const rightEyeCtrl = findFormControlByAliases(["visionrighteye", "right eye", "righteyevision", "vision - right eye"]);

    // Force a re-apply for the most common mandatory fields if they didn't stick.
    const weightVal = getRowValueByAliases(rowData, ["weightkg", "weight(kg)", "weight"]);
    const heightVal = getRowValueByAliases(rowData, ["heightlengthcm", "heightcm", "lengthcm", "height/length(cm)"]);
    const bpVal = getRowValueByAliases(rowData, ["bloodpressure", "bloodpressure*"]);
    const leftVal = getRowValueByAliases(rowData, ["visionlefteye", "vision-left-eye", "lefteyevision"]);
    const rightVal = getRowValueByAliases(rowData, ["visionrighteye", "vision-right-eye", "righteyevision"]);

    if (weightCtrl && String(weightVal || "").trim() && !String(weightCtrl.value || "").trim()) {
      setControlValue(weightCtrl, weightVal);
      weightCtrl.dispatchEvent(new Event("blur", { bubbles: true }));
    }
    if (heightCtrl && String(heightVal || "").trim() && !String(heightCtrl.value || "").trim()) {
      setControlValue(heightCtrl, heightVal);
      heightCtrl.dispatchEvent(new Event("blur", { bubbles: true }));
    }
    if (bpCtrl && String(bpVal || "").trim() && !String(bpCtrl.value || "").trim()) {
      setControlValue(bpCtrl, bpVal);
      bpCtrl.dispatchEvent(new Event("blur", { bubbles: true }));
    }
    if (leftEyeCtrl && String(leftVal || "").trim() && !String(leftEyeCtrl.value || "").trim()) {
      setControlValue(leftEyeCtrl, leftVal);
      leftEyeCtrl.dispatchEvent(new Event("blur", { bubbles: true }));
    }
    if (rightEyeCtrl && String(rightVal || "").trim() && !String(rightEyeCtrl.value || "").trim()) {
      setControlValue(rightEyeCtrl, rightVal);
      rightEyeCtrl.dispatchEvent(new Event("blur", { bubbles: true }));
    }

    // Wait for required controls to actually show values.
    const ready = await waitFor(() => {
      if (weightCtrl && !String(weightCtrl.value || "").trim()) return false;
      if (heightCtrl && !String(heightCtrl.value || "").trim()) return false;
      if (bpCtrl && !String(bpCtrl.value || "").trim()) return false;
      if (leftEyeCtrl && !String(leftEyeCtrl.value || "").trim()) return false;
      if (rightEyeCtrl && !String(rightEyeCtrl.value || "").trim()) return false;
      return true;
    }, timeoutMs, 150);

    // Some pages compute BMI after weight/height; if present and still empty, set it.
    if (bmiCtrl && !String(bmiCtrl.value || "").trim()) {
      const w = weightCtrl ? weightCtrl.value : weightVal;
      const h = heightCtrl ? heightCtrl.value : heightVal;
      const bmiComputed = calculateBmi(w, h);
      if (bmiComputed) {
        setControlValue(bmiCtrl, bmiComputed);
        bmiCtrl.dispatchEvent(new Event("blur", { bubbles: true }));
      }
    }

    // Give the UI a short moment to finalize validation/calculations.
    await sleep(350);
    return ready;
  }

  const DEFECT_TYPE_BUTTON_MAP = {
    defectsatbirth: "Defects at Birth",
    deficiencies: "Deficiencies",
    childhooddiseases: "Childhood Diseases",
    developmentdelay: "Development Delay",
    adolescenthealthconcerns: "Adolescent Health concerns"
  };

  function normalizeDefectType(raw) {
    return normalizeHeaderKey(String(raw || "").replace(/_/g, ""));
  }

  function groupRowsByPenId(rows) {
    const groups = new Map();
    for (const row of rows) {
      const penId = normalizePenId(getPenIdFromRowData(row));
      if (!penId) continue;
      if (!groups.has(penId)) groups.set(penId, []);
      groups.get(penId).push(row);
    }
    return groups;
  }

  function findScreeningHealthConditionsSection() {
    const headings = Array.from(document.querySelectorAll("h1, h2, h3, h4, h5, h6"));
    for (const heading of headings) {
      if (/screening\s*for\s*health\s*conditions/i.test(normalizeKey(heading.textContent || ""))) {
        return heading.closest("div, section, article") || heading.parentElement;
      }
    }
    return null;
  }

  function findScreeningCategoryButton(buttonText) {
    const target = normalizeHeaderKey(buttonText);

    const section = findScreeningHealthConditionsSection();
    const scope = section || document;

    const buttons = Array.from(
      scope.querySelectorAll("button[type='button'], button:not([type]), button")
    ).filter(isElementVisible);

    for (const btn of buttons) {
      if (normalizeHeaderKey(normalizeKey(btn.textContent || "")) === target) return btn;
    }
    for (const btn of buttons) {
      const t = normalizeHeaderKey(normalizeKey(btn.textContent || ""));
      if (t.includes(target) || target.includes(t)) return btn;
    }

    if (section) {
      const allButtons = Array.from(
        document.querySelectorAll("button[type='button'], button:not([type]), button")
      ).filter(isElementVisible);
      for (const btn of allButtons) {
        if (normalizeHeaderKey(normalizeKey(btn.textContent || "")) === target) return btn;
      }
    }

    return null;
  }

  function findHealthConditionsTable() {
    const section = findScreeningHealthConditionsSection();
    const scope = section || document;

    const h3s = Array.from(scope.querySelectorAll("h3"));
    const healthConditionsH3 = h3s.find((el) =>
      /^health\s*conditions$/i.test(normalizeKey(el.textContent || ""))
    );
    if (healthConditionsH3) {
      const container = healthConditionsH3.parentElement;
      if (container) {
        const table =
          healthConditionsH3.nextElementSibling?.tagName === "TABLE"
            ? healthConditionsH3.nextElementSibling
            : container.querySelector("table");
        if (table && tableContainsCheckboxRows(table)) {
          return {
            table,
            normalizedHeaders: ["select", "code", "healthconditionname"],
            useFixedIndices: true
          };
        }
      }
    }

    const tables = Array.from(scope.querySelectorAll("table"));
    for (const table of tables) {
      if (!tableContainsCheckboxRows(table)) continue;
      const headers = readHeadersFromTable(table);
      if (!headers.length) continue;
      const norm = headers.map((h) => normalizeHeaderKey(h));
      const hasSelect = norm.some((h) => h === "select" || h.startsWith("select"));
      const hasCode = norm.some((h) => h === "code");
      const hasConditionName = norm.some(
        (h) => h.includes("healthcondition") || (h.includes("condition") && h.includes("name"))
      );
      if (hasSelect && hasCode && hasConditionName) {
        return { table, normalizedHeaders: norm, useFixedIndices: false };
      }
    }
    for (const table of tables) {
      if (!tableContainsCheckboxRows(table)) continue;
      const headers = readHeadersFromTable(table);
      if (!headers.length) continue;
      const norm = headers.map((h) => normalizeHeaderKey(h));
      if (norm.some((h) => h === "code") && norm.some((h) => h.includes("condition") || h.includes("name"))) {
        return { table, normalizedHeaders: norm, useFixedIndices: false };
      }
    }
    return null;
  }

  function tableContainsCheckboxRows(table) {
    const tbody = table.querySelector("tbody");
    const rows = tbody ? table.querySelectorAll("tbody tr") : table.querySelectorAll("tr");
    for (const row of rows) {
      const firstTd = row.querySelector("td");
      if (firstTd && firstTd.querySelector("input[type='checkbox']")) return true;
    }
    return false;
  }

  async function waitForHealthConditionsTablePopulated(timeoutMs = 4000) {
    return waitFor(() => {
      const info = findHealthConditionsTable();
      if (!info) return false;
      const tbody = info.table.querySelector("tbody");
      const rows = tbody ? info.table.querySelectorAll("tbody tr") : info.table.querySelectorAll("tr");
      for (const row of rows) {
        const cells = row.querySelectorAll("td");
        if (cells.length >= 3) return true;
      }
      return false;
    }, timeoutMs, 150);
  }


  function getHealthConditionsTableElement() {
    const allH3 = Array.from(document.querySelectorAll("h3"));
    for (const h3 of allH3) {
      const text = (h3.textContent || "").replace(/\s+/g, " ").trim();
      if (/^health\s*conditions$/i.test(text)) {
        const parent = h3.parentElement;
        if (!parent) continue;
        const table = parent.querySelector("table");
        if (table) return table;
      }
    }
    return null;
  }

  async function checkHealthConditionByCode(code, defectNameForFallback) {
    const table = getHealthConditionsTableElement();
    if (!table) {
      console.warn("[HealthConditions] Health Conditions table NOT found on page.");
      return false;
    }

    const targetCode = String(code || "").trim();
    const rawName = String(defectNameForFallback || "")
      .replace(/^\d+\.?\s*/, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();

    let rows = Array.from(table.querySelectorAll("tbody tr"));
    if (!rows.length) {
      rows = Array.from(table.querySelectorAll("tr")).filter(
        (tr) => tr.querySelectorAll("td").length > 0
      );
    }

    console.warn(
      "[HealthConditions] Scanning", rows.length,
      "rows | target code:", JSON.stringify(targetCode),
      "| target name:", JSON.stringify(rawName)
    );

    for (const tr of rows) {
      const tds = Array.from(tr.querySelectorAll("td"));
      if (tds.length < 2) continue;

      const cb = tds[0]?.querySelector("input[type='checkbox']");
      if (!cb) continue;

      const codeText = (tds[1]?.textContent || "").trim();
      const nameText = tds.length >= 3
        ? (tds[2]?.textContent || "").replace(/\s+/g, " ").trim().toLowerCase()
        : "";

      const matchByCode = targetCode !== "" && codeText === targetCode;
      const matchByName =
        rawName.length > 0 &&
        (nameText === rawName ||
          nameText.includes(rawName) ||
          rawName.includes(nameText));

      if (matchByCode || matchByName) {
        const matchReason = matchByCode
          ? "CODE (" + codeText + ")"
          : 'NAME ("' + nameText + '")';
        tr.scrollIntoView({ block: "center", behavior: "instant" });
        await sleep(100);

        console.warn(
          "[HealthConditions] MATCH by", matchReason,
          "| row code:", codeText,
          "| row name:", (tds[2]?.textContent || "").trim(),
          "| checkbox.checked BEFORE:", cb.checked
        );

        if (cb.checked) {
          console.warn("[HealthConditions] Already checked — skipping click.");
          return true;
        }

        // Strategy 1: Native .click()
        cb.click();
        await sleep(200);
        if (cb.checked) {
          console.warn("[HealthConditions] Strategy 1 (click) SUCCESS — checked:", cb.checked);
          return true;
        }
        console.warn("[HealthConditions] Strategy 1 (click) did not stick, trying Strategy 2...");

        // Strategy 2: Set checked via native property descriptor + dispatch events
        const nativeSetter = Object.getOwnPropertyDescriptor(
          window.HTMLInputElement.prototype,
          "checked"
        )?.set;
        if (nativeSetter) {
          nativeSetter.call(cb, true);
        } else {
          cb.checked = true;
        }
        cb.dispatchEvent(new Event("input", { bubbles: true }));
        cb.dispatchEvent(new Event("change", { bubbles: true }));
        await sleep(200);
        if (cb.checked) {
          console.warn("[HealthConditions] Strategy 2 (setter+events) SUCCESS — checked:", cb.checked);
          return true;
        }
        console.warn("[HealthConditions] Strategy 2 did not stick, trying Strategy 3...");

        // Strategy 3: Full pointer/mouse event sequence (no native .click to avoid double-toggle)
        const evtOpts = { bubbles: true, cancelable: true, composed: true, view: window };
        for (const evtType of ["pointerdown", "mousedown", "pointerup", "mouseup", "click"]) {
          cb.dispatchEvent(new MouseEvent(evtType, evtOpts));
        }
        await sleep(200);
        if (cb.checked) {
          console.warn("[HealthConditions] Strategy 3 (mouseEvents) SUCCESS — checked:", cb.checked);
          return true;
        }
        console.warn("[HealthConditions] Strategy 3 did not stick, trying Strategy 4...");

        // Strategy 4: Click the parent <td> or <label>
        const parentTd = cb.closest("td");
        if (parentTd) {
          parentTd.click();
          await sleep(200);
        }
        if (cb.checked) {
          console.warn("[HealthConditions] Strategy 4 (parent td click) SUCCESS — checked:", cb.checked);
          return true;
        }

        // Strategy 5: Force checked + Angular-style ngModelChange via InputEvent
        if (nativeSetter) nativeSetter.call(cb, true);
        else cb.checked = true;
        cb.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true }));
        cb.dispatchEvent(new Event("change", { bubbles: true }));
        cb.dispatchEvent(new Event("blur", { bubbles: true }));
        await sleep(200);

        console.warn("[HealthConditions] FINAL checkbox.checked:", cb.checked);
        return cb.checked;
      }
    }

    console.warn(
      "[HealthConditions] NO MATCH | target code:",
      JSON.stringify(targetCode),
      "| target name:",
      JSON.stringify(rawName)
    );
    console.warn("[HealthConditions] Available rows in Health Conditions table:");
    for (const tr of rows) {
      const tds = Array.from(tr.querySelectorAll("td"));
      if (tds.length >= 3) {
        console.warn(
          "  code:", (tds[1]?.textContent || "").trim(),
          "| name:", (tds[2]?.textContent || "").trim()
        );
      }
    }
    return false;
  }

  function parseDefectCode(defectName) {
    const raw = String(defectName || "").trim();
    const leadingMatch = raw.match(/^(\d+)/);
    if (leadingMatch) return leadingMatch[1];
    const anyMatch = raw.match(/(\d+)/);
    return anyMatch ? anyMatch[1] : "";
  }

  async function handleOtherHealthCondition(defectName) {
    let container = document.querySelector(".other-defect-container");
    if (!container) {
      const allLabels = Array.from(document.querySelectorAll("label"));
      const label = allLabels.find((l) =>
        /other\s*health\s*condition/i.test(normalizeKey(l.textContent || ""))
      );
      container = label?.closest("div[class*='border']");
    }
    if (!container) {
      console.warn("[OtherHealth] Container not found.");
      return false;
    }

    container.scrollIntoView({ block: "center", behavior: "instant" });
    await sleep(200);

    const checkbox = container.querySelector("input[type='checkbox']");
    if (checkbox && !checkbox.checked) {
      dispatchClickSequence(checkbox);
      await sleep(600);
    }

    const defectLower = String(defectName || "").toLowerCase().trim();
    let radioValue = null;
    if (defectLower.includes("refer")) radioValue = "2";
    else if (defectLower.includes("on-site") || defectLower.includes("onsite") || defectLower.includes("on site"))
      radioValue = "1";

    if (radioValue) {
      const radio = container.querySelector(
        `input[type='radio'][name='otherHealthConditionOption'][value='${radioValue}']`
      );
      if (radio && !radio.checked) {
        dispatchClickSequence(radio);
        await sleep(400);
      }
    }

    const textarea = container.querySelector("textarea");
    if (textarea) {
      textarea.scrollIntoView({ block: "center", behavior: "instant" });
      textarea.focus();
      await sleep(100);
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value")?.set;
      if (setter) setter.call(textarea, "...");
      else textarea.value = "...";
      textarea.dispatchEvent(new Event("input", { bubbles: true }));
      textarea.dispatchEvent(new Event("change", { bubbles: true }));
    }

    console.warn("[OtherHealth] Completed:", defectName, "radio:", radioValue);
    return true;
  }

  async function waitForHealthConditionsSection(timeoutMs = 8000) {
    return waitFor(() => Boolean(findScreeningHealthConditionsSection()), timeoutMs, 150);
  }

  // ── Referral Details ──

  async function selectReactSelectOption(container, targetText) {
    const control = container.querySelector(".select__control, [class*='control']");
    if (!control) {
      console.warn("[Referral] react-select control not found in container.");
      return false;
    }

    control.scrollIntoView({ block: "center", behavior: "instant" });
    dispatchClickSequence(control);
    await sleep(400);

    const input = container.querySelector("input[type='text'], input[aria-autocomplete]");
    if (input) {
      input.focus();
      const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value")?.set;
      if (nativeSetter) nativeSetter.call(input, targetText);
      else input.value = targetText;
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
      await sleep(600);
    }

    const normalizedTarget = targetText.toLowerCase().replace(/\s+/g, " ").trim();
    let options = Array.from(container.querySelectorAll(
      "[class*='option'], [class*='menu'] div[id*='option'], [role='option']"
    ));

    if (!options.length) {
      options = findOpenReactSelectMenu();
    }

    let match = options.find((opt) => {
      const text = (opt.textContent || "").replace(/\s+/g, " ").trim().toLowerCase();
      return text === normalizedTarget || text.includes(normalizedTarget) || normalizedTarget.includes(text);
    });

    if (!match && options.length > 0) {
      match = options[0];
    }

    if (match) {
      match.scrollIntoView({ block: "center", behavior: "instant" });
      dispatchClickSequence(match);
      await sleep(300);
      console.warn("[Referral] Selected option:", match.textContent?.trim());
      return true;
    }

    console.warn("[Referral] No matching option found for:", targetText,
      "| available:", options.map((o) => o.textContent?.trim()));
    return false;
  }

  function findOpenReactSelectMenu() {
    const selectors = [
      "[class*='menu'] [class*='option']",
      "[class*='MenuList'] [class*='option']",
      "[class*='menu-list'] [class*='option']",
      "div[id*='option']",
      "[role='listbox'] [role='option']"
    ];
    for (const sel of selectors) {
      const opts = document.querySelectorAll(sel);
      if (opts.length > 0) return Array.from(opts);
    }
    return [];
  }

  async function selectFirstReactSelectOption(container) {
    const control = container.querySelector(".select__control, [class*='control']");
    if (!control) {
      console.warn("[Referral] react-select control not found in container.");
      return false;
    }

    control.scrollIntoView({ block: "center", behavior: "instant" });
    dispatchClickSequence(control);
    await sleep(1000);

    let options = Array.from(container.querySelectorAll(
      "[class*='option'], [class*='menu'] div[id*='option'], [role='option']"
    ));

    if (!options.length) {
      options = findOpenReactSelectMenu();
    }

    if (options.length > 0) {
      dispatchClickSequence(options[0]);
      await sleep(300);
      console.warn("[Referral] Selected first option:", options[0].textContent?.trim());
      return true;
    }

    console.warn("[Referral] No options found, retrying with keyboard...");
    const input = container.querySelector("input") || control.querySelector("input");
    if (input) {
      input.focus();
      await sleep(300);
      input.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", keyCode: 40, bubbles: true }));
      await sleep(400);
      input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", keyCode: 13, bubbles: true }));
      await sleep(300);
      console.warn("[Referral] Used keyboard to select first option.");
      return true;
    }

    console.warn("[Referral] No options found in Facility Name dropdown.");
    return false;
  }

  function findLabeledReactSelect(labelText) {
    const labelRegex = new RegExp(labelText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");
    const labels = Array.from(document.querySelectorAll("label"));
    for (const label of labels) {
      const text = (label.textContent || "").replace(/\s+/g, " ").trim();
      if (labelRegex.test(text)) {
        const parent = label.closest("div");
        if (parent) {
          const selectContainer = parent.querySelector("[class*='select__control'], [class*='css-']")?.closest("[class*='container'], .basic-single, [class*='select']");
          if (selectContainer) return selectContainer;
          const reactSelect = parent.querySelector("[class*='select']");
          if (reactSelect) return reactSelect;
        }
      }
    }

    const allSelects = Array.from(document.querySelectorAll("[class*='select__control']"));
    for (const ctrl of allSelects) {
      const wrapper = ctrl.closest("div.grid, div[class*='grid']");
      if (wrapper) {
        const lbl = wrapper.querySelector("label");
        if (lbl && labelRegex.test((lbl.textContent || "").replace(/\s+/g, " ").trim())) {
          return ctrl.closest("[class*='container'], .basic-single") || ctrl.parentElement;
        }
      }
    }
    return null;
  }

  async function fillReferralDetails(referalValue) {
    console.warn("[Referral] Starting referral fill. Excel value:", referalValue);

    const checkbox = document.getElementById("referSameInstitute")
      || document.querySelector("input[type='checkbox'][id*='refer']");
    if (checkbox) {
      checkbox.scrollIntoView({ block: "center", behavior: "instant" });
      await sleep(100);
      if (!checkbox.checked) {
        checkbox.click();
        await sleep(200);
        if (!checkbox.checked) {
          const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "checked")?.set;
          if (setter) setter.call(checkbox, true);
          else checkbox.checked = true;
          checkbox.dispatchEvent(new Event("change", { bubbles: true }));
          checkbox.dispatchEvent(new Event("input", { bubbles: true }));
        }
        await sleep(500);
      }
      console.warn("[Referral] 'Refer to same facility' checkbox checked:", checkbox.checked);
    } else {
      console.warn("[Referral] 'Refer to same facility' checkbox not found.");
    }

    if (referalValue) {
      const facilityTypeContainer = findLabeledReactSelect("Facility Type");
      if (facilityTypeContainer) {
        console.warn("[Referral] Found Facility Type dropdown, selecting:", referalValue);
        await selectReactSelectOption(facilityTypeContainer, referalValue);
        await sleep(500);
      } else {
        console.warn("[Referral] Facility Type dropdown not found on page.");
      }
    }

    const facilityNameContainer = findLabeledReactSelect("Facility Name");
    if (facilityNameContainer) {
      console.warn("[Referral] Found Facility Name dropdown, selecting first option...");
      await selectFirstReactSelectOption(facilityNameContainer);
      await sleep(500);
    } else {
      console.warn("[Referral] Facility Name dropdown not found on page.");
    }

    console.warn("[Referral] Referral details fill complete.");
    return true;
  }

  async function submitWithMobileNumber(mobileNumber) {
    console.warn("[Submit] Starting Preview & Submit flow. Mobile from Excel:", mobileNumber || "(none)");

    const previewBtn = findPreviewSubmitButton();
    if (!previewBtn) {
      console.warn("[Submit] Preview & Submit button not found.");
      return false;
    }
    previewBtn.scrollIntoView({ block: "center", behavior: "instant" });
    await sleep(150);
    // Use a single click to avoid duplicate submit/modal handlers.
    clickElementOnce(previewBtn);
    console.warn("[Submit] Preview & Submit clicked. Waiting for modal...");

    // Wait for modal: either Confirm button or Mobile input (phone-exists path may skip mobile step)
    const modalReady = await waitFor(() => {
      const confirmBtn = findConfirmSubmitButton();
      if (confirmBtn) return true;
      const inputs = document.querySelectorAll("input[placeholder*='Mobile'], input[placeholder*='mobile'], input[placeholder*='Enter Mobile']");
      return inputs.length > 0;
    }, 6000, 120);

    if (!modalReady) {
      console.warn("[Submit] Modal (Confirm or Mobile input) did not appear.");
      return false;
    }
    await sleep(300);

    const mobileInput = document.querySelector(
      "input[placeholder*='Mobile'], input[placeholder*='mobile'], input[placeholder*='Enter Mobile']"
    );

    const hasExistingPhone = mobileInput && String(mobileInput.value || "").trim().length > 0;
    if (hasExistingPhone) {
      console.warn("[Submit] Phone number already exists on page. Skipping mobile input, clicking Confirm.");
    } else if (mobileInput && mobileNumber) {
      mobileInput.scrollIntoView({ block: "center", behavior: "instant" });
      mobileInput.focus();
      await sleep(80);
      const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value")?.set;
      if (nativeSetter) nativeSetter.call(mobileInput, mobileNumber);
      else mobileInput.value = mobileNumber;
      mobileInput.dispatchEvent(new Event("input", { bubbles: true }));
      mobileInput.dispatchEvent(new Event("change", { bubbles: true }));
      await sleep(200);
      console.warn("[Submit] Mobile number entered from Excel:", mobileNumber);
    } else if (mobileInput && !mobileNumber) {
      console.warn("[Submit] Mobile input is empty and no mobile number in Excel. Proceeding to Confirm.");
    }

    await sleep(200);
    const confirmBtn = findConfirmSubmitButton();
    if (confirmBtn) {
      confirmBtn.scrollIntoView({ block: "center", behavior: "instant" });
      await sleep(200);
      // Use a single click to avoid double submission side-effects.
      clickElementOnce(confirmBtn);
      console.warn("[Submit] Confirm & Submit clicked.");
      await waitFor(() => !findConfirmSubmitButton(), 12000, 200);
      await sleep(1000);
      console.warn("[Submit] Submission complete.");
      return true;
    }

    console.warn("[Submit] Confirm & Submit button not found.");
    return false;
  }

  async function fillHealthConditionsFromRows(rows) {
    const results = { filled: 0, failed: 0, skipped: false, details: [] };

    // If Birth Defect Found = No, do not enforce/attempt defect entry.
    const birthDefectFoundRaw = rows?.length
      ? getRowValueByAliases(rows[0], ["birthdefectfound", "birth defect found", "birthdefect"])
      : "";
    const birthDefectFoundNorm = normalizeHeaderKey(String(birthDefectFoundRaw || ""));
    if (birthDefectFoundNorm === "no") {
      results.skipped = true;
      results.details.push({ type: "Birth Defect Found", status: "skipped_defect_entry_because_no" });
      console.warn("[HealthConditions] Birth Defect Found = No. Skipping defect/health-conditions selection.");
      return results;
    }

    const byType = {};
    for (const row of rows) {
      const defectType = getRowValueByAliases(row, ["selectdefecttype", "defecttype"]);
      if (!defectType) continue;
      const key = normalizeDefectType(defectType);
      if (!byType[key]) byType[key] = { original: defectType, rows: [] };
      byType[key].rows.push(row);
    }

    if (!Object.keys(byType).length) {
      console.warn("[HealthConditions] No defect types found in row data. Row keys:", rows.length ? Object.keys(rows[0]) : "no rows");
      return results;
    }

    const section = findScreeningHealthConditionsSection();
    if (!section) {
      console.warn("[HealthConditions] Screening for Health Conditions section not found, waiting...");
      const sectionReady = await waitForHealthConditionsSection(15000);
      if (!sectionReady) {
        console.warn("[HealthConditions] Section still not found after waiting. Proceeding anyway.");
      }
    }

    const sectionEl = findScreeningHealthConditionsSection();
    if (sectionEl) {
      sectionEl.scrollIntoView({ block: "start", behavior: "instant" });
      await sleep(300);
    }

    for (const [typeKey, group] of Object.entries(byType)) {
      console.warn("[HealthConditions] Processing defect type:", group.original, "→ normalized:", typeKey);

      if (typeKey === "otherhealthcondition") {
        for (const row of group.rows) {
          const defectName = getRowValueByAliases(row, ["defectname", "defectother"]);
          try {
            const ok = await handleOtherHealthCondition(defectName);
            results[ok ? "filled" : "failed"]++;
            results.details.push({
              type: group.original,
              name: defectName,
              status: ok ? "ok" : "container_not_found"
            });
          } catch (err) {
            results.failed++;
            results.details.push({
              type: group.original,
              name: defectName,
              status: err instanceof Error ? err.message : String(err)
            });
          }
        }
        continue;
      }

      const buttonText = DEFECT_TYPE_BUTTON_MAP[typeKey];
      if (!buttonText) {
        console.warn("[HealthConditions] Unknown defect type key:", typeKey, "original:", group.original);
        for (const row of group.rows) {
          results.failed++;
          results.details.push({ type: group.original, status: "unknown_defect_type" });
        }
        continue;
      }

      const btn = findScreeningCategoryButton(buttonText);
      if (!btn) {
        console.warn("[HealthConditions] Button not found for:", buttonText);
        for (const row of group.rows) {
          results.failed++;
          results.details.push({ type: group.original, status: "button_not_found: " + buttonText });
        }
        continue;
      }

      console.warn("[HealthConditions] Clicking category button:", buttonText, btn.textContent);
      dispatchClickSequence(btn);
      await sleep(800);

      const tableReady = await waitForHealthConditionsTablePopulated(12000);
      if (!tableReady) {
        console.warn("[HealthConditions] Table did not populate in time after clicking", buttonText);
      }
      await sleep(400);

      for (const row of group.rows) {

        const defectName = getRowValueByAliases(row, ["defectname"]);
        const defectOther = getRowValueByAliases(row, ["defectother"]);
        const identCode = getRowValueByAliases(row, ["identificationcode", "idcode"]);

        // Fallback: try direct key access if aliases returned empty
        const directDefectName = defectName
          || row["Defect Name"] || row["Defect_Name"] || row["DefectName"] || "";

        const codeFromName = parseDefectCode(directDefectName);
        const codeFromOther = parseDefectCode(defectOther);
        const codeFromIdent = parseDefectCode(identCode);
        const bestCode = codeFromName || codeFromOther || codeFromIdent;

        const nameFromDefect = String(directDefectName || "").replace(/^\d+\.?\s*/, "").trim();
        const nameFromOther = String(defectOther || "").replace(/^\d+\.?\s*/, "").trim();
        const bestName = nameFromDefect || nameFromOther;

        console.warn("[HealthConditions] Resolved → defectName:", JSON.stringify(directDefectName),
          "| bestCode:", JSON.stringify(bestCode), "| bestName:", JSON.stringify(bestName));

        const ok = await checkHealthConditionByCode(bestCode, bestName);
        console.warn(
          "[HealthConditions] Result for code:", JSON.stringify(bestCode || ""),
          "name:", JSON.stringify(bestName),
          "→", ok ? "SELECTED" : "NOT FOUND"
        );
        results[ok ? "filled" : "failed"]++;
        results.details.push({
          type: group.original,
          name: defectName || defectOther,
          code: bestCode || "",
          status: ok ? "ok" : "not_found"
        });
        await sleep(200);
      }
    }

    return results;
  }

  async function autoFillRowsOnPage(rows, options = {}) {
    const stopAfterFirst = options?.stopAfterFirst !== false;
    const autoSubmit = options?.autoSubmit !== false;
    const summary = {
      total: rows.length,
      attempted: 0,
      success: 0,
      failed: 0,
      failures: []
    };

    if (!isChildScreeningPage()) {
      throw new Error("Open the child screening table page first: /RBSK/childScreening");
    }
    const baseChildScreeningUrl = location.href;
    const penIdGroups = groupRowsByPenId(rows);

    // Ensure table is ready, then start from first page so the first PEN ID is found immediately.
    const tableReady = await waitFor(() => Boolean(getBestStudentTableContext()), 3000, 100);
    if (!tableReady) throw new Error("Student table not visible. Wait for the page to load and try again.");
    await goToFirstTablePage();

    for (const [penId, groupRows] of penIdGroups) {
      summary.attempted += 1;

      try {
        const rowMatch = await findRowByPenIdAcrossPages(penId);
        if (!rowMatch?.liveRow) {
          summary.failed += 1;
          summary.failures.push({
            penId,
            reason: `No matching row found on any page for PEN ID / AWC Child ID "${penId}". Skipping.`
          });
          continue;
        }
        const { liveRow, actionIndex } = rowMatch;

        const actionEl = findScreeningStartElementInRow(liveRow, actionIndex);
        if (!actionEl) throw new Error("Screening Start button not found in this row.");

        await openScreeningDetailsForRow(liveRow, actionEl, { useLegacyClick: false });
        assertDetailsPenIdMatches(penId);
        const fillInfo = fillScreeningDetailsFormFromRow(groupRows[0]);
        await settleMandatoryScreeningDetails(groupRows[0], 15000);
        console.warn("[AutoFill] PEN:", penId, "| rows:", groupRows.length, "| defectTypes:", groupRows.map(r => r["Select Defect type"] || r["Select_Defect_type"] || r["Defect type"] || "(none)"));

        // Ensure Birth Defect Found is applied before defect logic.
        const birthDefectFound = getRowValueByAliases(groupRows[0], ["birthdefectfound", "birth defect found", "birthdefect"]);
        const birthNorm = normalizeHeaderKey(String(birthDefectFound || ""));
        if (birthNorm) {
          const ctrl = findFormControlByAliases(["birthdefectfound", "birth defect found", "birth defect"]);
          if (ctrl) {
            // wait until the UI reflects the selected value before proceeding
            await waitFor(() => normalizeHeaderKey(String(ctrl.value || ctrl.textContent || "")) === birthNorm, 6000, 150);
            await sleep(300);
          }
        }

        const healthResults = await fillHealthConditionsFromRows(groupRows);
        console.warn("[AutoFill] Health conditions result:", JSON.stringify(healthResults));

        const referalValue = getRowValueByAliases(groupRows[0], ["referal", "referral", "referaltype", "referraltype"]);
        console.warn("[AutoFill] Referral value from Excel:", JSON.stringify(referalValue));
        if (referalValue) {
          await fillReferralDetails(referalValue);
        }

        const mobileNumber = getRowValueByAliases(groupRows[0], ["mobilenumber", "mobile", "phonenumber", "phone"]);
        const submitted = await submitWithMobileNumber(mobileNumber);
        console.warn("[AutoFill] Submit result:", submitted);

        if (submitted) {
          summary.success += 1;
          await returnToStudentTable(baseChildScreeningUrl);
          await sleep(1000);
        } else {
          summary.failed += 1;
          summary.failures.push({ penId, reason: "Submit failed (Preview/Confirm step)" });
          try { await returnToStudentTable(baseChildScreeningUrl); } catch (_e) {}
          await sleep(500);
        }
        console.warn("[AutoFill] === PEN", penId, "DONE ===", summary.success, "/", penIdGroups.size, "completed,", summary.failed, "failed");

        if (stopAfterFirst) {
          return {
            ...summary,
            pausedOnDetails: true,
            currentPenId: penId,
            filledCount: fillInfo.filledCount,
            missingFields: fillInfo.missingFields,
            healthConditions: healthResults
          };
        }
      } catch (error) {
        summary.failed += 1;
        summary.failures.push({
          penId,
          reason: error instanceof Error ? error.message : String(error)
        });
        try {
          await returnToStudentTable(baseChildScreeningUrl);
        } catch (_returnErr) {
          // Continue to next PEN ID group.
        }
      }
    }

    return summary;
  }

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (message?.type === "SCRAPE_PAGE_DATA") {
      try {
        const penId = extractPenId();
        const fields = scrapeFormFields();
        sendResponse({
          ok: true,
          data: {
            url: location.href,
            title: document.title,
            penId,
            fields
          }
        });
      } catch (error) {
        sendResponse({
          ok: false,
          error: error instanceof Error ? error.message : String(error)
        });
      }
      return;
    }

    if (message?.type === "AUTO_FILL_EXCEL_ROWS") {
      (async () => {
        try {
          const rows = Array.isArray(message.rows) ? message.rows : [];
          const summary = await autoFillRowsOnPage(rows, message.options || {});
          sendResponse({ ok: true, summary });
        } catch (error) {
          sendResponse({
            ok: false,
            error: error instanceof Error ? error.message : String(error)
          });
        }
      })();
      return true;
    }

    if (message?.type === "EXTRACT_STUDENT_TABLE") {
      const opts = message?.options || {};
      if (opts && opts.allPages) {
        (async () => {
          try {
            const data = await extractStudentTableDataAllPages(opts);
            sendResponse({
              ok: true,
              data
            });
          } catch (error) {
            sendResponse({
              ok: false,
              error: error instanceof Error ? error.message : String(error)
            });
          }
        })();
        return true;
      }

      try {
        const data = extractStudentTableData();
        sendResponse({
          ok: true,
          data
        });
      } catch (error) {
        sendResponse({
          ok: false,
          error: error instanceof Error ? error.message : String(error)
        });
      }
      return;
    }

    if (message?.type === "EXTRACT_STUDENT_TABLE_WITH_DETAILS") {
      (async () => {
        try {
          const data = await extractStudentTableWithDetails(message.options || {});
          sendResponse({
            ok: true,
            data
          });
        } catch (error) {
          sendResponse({
            ok: false,
            error: error instanceof Error ? error.message : String(error)
          });
        }
      })();
      return true;
    }
  });
})();
