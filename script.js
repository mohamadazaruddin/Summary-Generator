(() => {
  const el = (id) => document.getElementById(id);
  const fileEl = el("file");
  const statusEl = el("status");
  const btnGenerate = el("btnGenerate");
  const btnCopy = el("btnCopy");
  const btnClear = el("btnClear");
  const summaryEditor = el("summaryEditor");
  const focalTypeEl = document.getElementById("focalType");
  const individualFormEl = document.getElementById("focal-individual-form");
  const entityFormEl = document.getElementById("focal-entity-form");
  let rows = [];

  // function for all the the transaction data points
  function TransactionDataPoints() {
    const totalAmount = calculatingTotal(rows, "AMOUNT");
    const totalRowCount = calculateCount(rows, "AMOUNT", {
      countNonEmpty: false,
    });
    const { start, end } = calculatedateRange(rows, "Transaction Date");
    const result = caluculatevalue(rows, "Transaction Type", "p2p");
    console.log(
      totalAmount,
      "rows",
      totalRowCount,
      start,
      "-",
      end,
      result,
      rows,
    );
  }

  function calculatingTotal(rows, headerName) {
    if (!Array.isArray(rows)) throw new TypeError("rows must be an array");
    if (!headerName) return 0;

    const target = String(headerName).trim().toLowerCase();

    // Find the actual key in the data (case/space-insensitive)
    const keys = rows.flatMap((r) =>
      r && typeof r === "object" ? Object.keys(r) : [],
    );
    const key =
      keys.find((k) => String(k).trim().toLowerCase() === target) || headerName;

    const toNumber = (v) => {
      if (typeof v === "number") return v;
      if (typeof v !== "string") return NaN;
      const cleaned = v.replace(/[$,]/g, "").trim();
      return cleaned ? Number(cleaned) : NaN;
    };

    return rows.reduce((sum, r) => {
      const n = toNumber(r?.[key]);
      return Number.isFinite(n) ? sum + n : sum;
    }, 0);
  }

  function calculateCount(rows, headerName, { countNonEmpty = true } = {}) {
    if (!Array.isArray(rows)) throw new TypeError("rows must be an array");
    if (!headerName) return 0;

    const target = String(headerName).trim().toLowerCase();

    // Find the actual key in the data (case/space-insensitive)
    const keys = rows.flatMap((r) =>
      r && typeof r === "object" ? Object.keys(r) : [],
    );
    const key =
      keys.find((k) => String(k).trim().toLowerCase() === target) || headerName;

    if (!countNonEmpty) return rows.length;

    return rows.reduce((count, r) => {
      const v = r?.[key];
      const nonEmpty =
        v !== null &&
        v !== undefined &&
        !(typeof v === "string" && v.trim() === "");
      return nonEmpty ? count + 1 : count;
    }, 0);
  }
  function caluculatevalue(rows, headerName, type) {
    if (!Array.isArray(rows)) throw new TypeError("rows must be an array");
    if (!headerName || !type) {
      return { debitTotal: 0, creditTotal: 0, total: 0 };
    }

    const normalize = (v) =>
      String(v || "")
        .trim()
        .toLowerCase();

    const findKey = (rows, targetHeader) => {
      const target = normalize(targetHeader);
      const keys = rows.flatMap((r) =>
        r && typeof r === "object" ? Object.keys(r) : [],
      );
      return keys.find((k) => normalize(k) === target) || targetHeader;
    };

    const typeKey = findKey(rows, headerName); // e.g. TRANSACTION TYPE
    const amountKey = findKey(rows, "AMOUNT"); // AMOUNT
    const creditDebitKey = findKey(rows, "CREDIT / DEBIT"); // CREDIT / DEBIT

    const toNumber = (v) => {
      if (typeof v === "number") return v;
      if (typeof v !== "string") return NaN;
      const cleaned = v.replace(/[$,]/g, "").trim();
      return cleaned ? Number(cleaned) : NaN;
    };

    let debitTotal = 0;
    let creditTotal = 0;

    for (const row of rows) {
      if (normalize(row?.[typeKey]) !== normalize(type)) continue;

      const amount = toNumber(row?.[amountKey]);
      const direction = normalize(row?.[creditDebitKey]);

      if (!Number.isFinite(amount)) continue;

      if (direction === "debit") debitTotal += amount;
      if (direction === "credit") creditTotal += amount;
    }

    return {
      debitTotal,
      creditTotal,
      total: debitTotal + creditTotal,
    };
  }

  const formatMMDDYYYY = (d) => {
    if (!(d instanceof Date) || Number.isNaN(d.getTime())) return null;
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const yyyy = d.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
  };

  function calculatedateRange(rows, dateHeaderName) {
    if (!Array.isArray(rows)) throw new TypeError("rows must be an array");
    if (!dateHeaderName) return { start: null, end: null, key: null };

    const target = String(dateHeaderName).trim().toLowerCase();

    const keys = rows.flatMap((r) =>
      r && typeof r === "object" ? Object.keys(r) : [],
    );

    const key =
      keys.find((k) => String(k).trim().toLowerCase() === target) ||
      dateHeaderName;

    const toDate = (v) => {
      if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v;
      if (typeof v === "number") {
        const d = new Date(v);
        return Number.isNaN(d.getTime()) ? null : d;
      }
      if (typeof v !== "string") return null;

      const s = v.trim();
      if (!s) return null;

      // Strict MM/DD/YYYY only
      const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (!m) return null;

      const mm = Number(m[1]);
      const dd = Number(m[2]);
      const yyyy = Number(m[3]);
      if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;

      const d = new Date(yyyy, mm - 1, dd);

      // Reject rollover dates (e.g., 02/31 -> Mar 2)
      if (
        d.getFullYear() !== yyyy ||
        d.getMonth() !== mm - 1 ||
        d.getDate() !== dd
      )
        return null;

      return d;
    };

    let start = null;
    let end = null;

    for (const r of rows) {
      const d = toDate(r?.[key]);
      if (!d) continue;
      if (!start || d < start) start = d;
      if (!end || d > end) end = d;
    }

    return { start: formatMMDDYYYY(start), end: formatMMDDYYYY(end), key };
  }
  function setStatus(msg, kind = "") {
    statusEl.textContent = msg || "";
    statusEl.className = "status" + (kind ? ` ${kind}` : "");
  }
  function fmtMoney(n) {
    if (!Number.isFinite(n)) return "N/A";
    const sign = n < 0 ? "-" : "";
    const abs = Math.abs(n);
    return `${sign}${abs.toLocaleString(undefined, {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`;
  }
  function normalize(s) {
    return String(s ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ")
      .replace(/[^a-z0-9 ]/g, "");
  }

  function findSheetByName(wb, wantedName) {
    const wanted = normalize(wantedName);
    return (wb.SheetNames || []).find((n) => normalize(n) === wanted) || "";
  }
  function setFocalFormVisibility() {
    const type = focalTypeEl?.value || "individual";
    const isIndividual = type === "individual";

    individualFormEl.classList.toggle("hidden", !isIndividual);
    entityFormEl.classList.toggle("hidden", isIndividual);
  }

  function esc(s) {
    return String(s ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function buildRfiSummary(focalEntityName) {
    const fe = (focalEntityName || "").trim() || "N/A";

    const status =
      document.getElementById("rfi_status")?.value || "no_prior_rfi_flag";
    const valOrNA = (id) =>
      (document.getElementById(id)?.value || "").trim() || "N/A";

    if (status === "no_prior_rfi_flag") {
      return `There are no prior RFIs identified for ${fe}.`;
    }

    if (status === "prior_rfi_flag_present") {
      const count = valOrNA("rfi_count");
      const submitted = valOrNA("rfi_submitted_date");
      const questions = valOrNA("rfi_questions_summary");
      const responseDate = valOrNA("rfi_questions_received_date");
      const responseDetails = valOrNA("rfi_response_detail");

      return `A total of ${count} prior RFIs were identified for ${fe}. An RFI was submitted by TD on ${submitted} to inquire about ${questions}. A response was received on ${responseDate}. The customer/branch stated: ${responseDetails}.`;
    }

    if (status === "lookback_submitted_no_response") {
      const submitted = valOrNA("rfi_lookback_submitted_date");
      const purpose = valOrNA("rfi_lookback_purpose");
      const today = valOrNA("rfi_todays_date");

      return `During the Lookback an RFI was submitted to TD on ${submitted} regarding ${purpose}. As of ${today}, no response has been received.`;
    }

    if (status === "lookback_submitted_response_received") {
      const submitted = valOrNA("rfi_lookback_resp_submitted_date");
      const purpose = valOrNA("rfi_lookback_resp_purpose");
      const received = valOrNA("rfi_lookback_r  esp_received_date");
      const summary = valOrNA("rfi_response_summary");

      return `During the Lookback, an RFI was submitted to TD on ${submitted} regarding ${purpose}. A response was received on ${received}, stating: ${summary}.`;
    }

    return "";
  }
  function wireCustomerStatusToggle() {
    const statusEl = document.getElementById("ind_customerStatus");
    const inactiveWrap = document.getElementById("ind_inactive_fields");
    const endDateEl = document.getElementById("ind_endDate");
    const reasonEl = document.getElementById("ind_reason");
    if (!statusEl || !inactiveWrap) return;

    const apply = () => {
      const isInactive = statusEl.value === "inactive";
      inactiveWrap.style.display = isInactive ? "" : "none";

      // Optional: clear values when hidden
      if (!isInactive) {
        if (endDateEl) endDateEl.value = "";
        if (reasonEl) reasonEl.value = "";
      }
    };

    statusEl.addEventListener("change", apply);
    apply(); // run once on load
  }
  function buildCounterpartySummary() {
    const count =
      Number.parseInt(
        document.getElementById("counterparty_count")?.value,
        10,
      ) || 0;

    const lines = [];

    for (let i = 1; i <= count; i++) {
      const name = document.getElementById(`cp_${i}_name`)?.value?.trim() || "";
      const reason =
        document.getElementById(`cp_${i}_reason`)?.value?.trim() || "";
      const addl = document.getElementById(`cp_${i}_addl`)?.value?.trim() || "";

      if (!name && !reason && !addl) continue;

      let s = `Counterparty ${name || `#${i}`} has been sampled`;
      if (reason) s += ` due to ${reason}`;
      s += `.`;
      if (addl) s += ` “${addl}”.`;

      lines.push(s);
    }

    if (lines.length === 0) {
      return ["No Counterparty is selected for sampling."];
    }

    return lines;
  }

  function buildAnalystDisposition() {
    const name = valOrNA("ind_focalName");
    const occupation = valOrNA("ind_occupation");
    const employer = valOrNA("ind_employer");
    return `
      
      <p>
      The Focal Entity, ${name}, is a ${occupation} at ${employer}. The Focal Entity maintains
        [ACCOUNT TYPES]. The credits on the account reflected [TYPES OF CREDITS, GENERAL DESCRIPTION
        OF ORIGINATOR(S), AND APPARENT PURPOSE OF THE CREDITS]. Account debits primarily reflect
        [GENERAL DESCRIPTION OF DEBITS AND THEIR APPARENT PURPOSE]. The activity appears reasonable
        [MITIGATING REASONS FOR THE ACTIVITY].
      </p>

      <p>
        During the review period of [REVIEW PERIOD DATE RANGE], review of transactional activity
        identified that the activity consisted of [GENERAL DESCRIPTION OF ACCOUNT ACTIVITY]. This appears
        consistent with [APPARENT PURPOSE OF ACTIVITY] (explain why the activity makes sense for the customer).
        The alerting activity consisted of [DESCRIPTION OF ALERTING ACTIVITY], which appears to be
        [APPARENT PURPOSE OF ALERTING TRANSACTIONS]. As for [MITIGATING FACTORS], this case is recommended
        for closure.
      </p>
 
  `.trim();
  }
  function p2psummary() {
    const { start, end } = calculatedateRange(rows, "Transaction Date");
    return `
      <p>
      Between ${start} - ${end}, the Focal Entity received [# P2P credits] P2P transactions through [list P2P Platform names] totaling [$ total P2P credit amount] ([percentage of total credits by value originating from P2P transactions]% of total credits) (if available: from [#] Counterparties) and conducted [# P2P debit] P2P transactions through [list P2P Platform names] totaling [$ total P2P debit amount] ([percentage of total debits by value originating from P2P transactions]% of total activity) (if available: from [#] Counterparties). [Include a short description about whether the P2P activity appears reasonable for the customer]. 
      </p>

  `.trim();
  }
  function cashSummary() {
    return `
      <p>
      Between [DATES OF FULL REVIEW PERIOD], the Focal Entity conducted [# cash deposits] cash deposits via [branch, ATM, armored car], at locations in [geographic location of deposits], totaling [$ total cash credit amount] ([percentage of total credits by value originating from cash transactions]% of total credits) and conducted e] cash withdrawals via [branch, ATM, armored car], at locations in [geographic location of deposits], totaling [$ total cash debit amount] ([percentage of total debits by value being withdrawn as cash]% of total debits) [Include a short description about whether the cash activity appears].
      </p>

  `.trim();
  }
  function Conclusion() {
    return `
      
      <p>
      The review of transactional activity identified no suspicious patterns and therefore this case is recommended for closure.
      </p>

  `.trim();
  }

  function wireRfiStatusFields() {
    const statusEl = document.getElementById("rfi_status");
    const priorEl = document.getElementById("rfi_fields_prior");
    const noRespEl = document.getElementById("rfi_fields_lookback_noresp");
    const respEl = document.getElementById("rfi_fields_lookback_resp");
    const todaysEl = document.getElementById("rfi_todays_date");

    if (!statusEl || !priorEl || !noRespEl || !respEl) return;

    const setTodayIfEmpty = () => {
      if (!todaysEl) return;
      if (!todaysEl.value)
        todaysEl.value = new Date().toISOString().slice(0, 10);
    };

    const sync = () => {
      const v = statusEl.value;

      priorEl.classList.toggle("hidden", v !== "prior_rfi_flag_present");
      noRespEl.classList.toggle(
        "hidden",
        v !== "lookback_submitted_no_response",
      );
      respEl.classList.toggle(
        "hidden",
        v !== "lookback_submitted_response_received",
      );

      if (v === "lookback_submitted_no_response") setTodayIfEmpty();
    };

    statusEl.addEventListener("change", sync);
    sync();
  }
  function wireCounterpartyAddButton() {
    const countEl = document.getElementById("counterparty_count");
    const btnEl = document.getElementById("counterparty_add_btn");
    const wrapEl = document.getElementById("counterparty_fields");
    if (!countEl || !btnEl || !wrapEl) return;

    const build = () => {
      wrapEl.innerHTML = "";

      const count = Math.max(
        0,
        Math.min(100, Number.parseInt(countEl.value, 10) || 0),
      ); // cap at 100
      for (let i = 1; i <= count; i++) {
        const section = document.createElement("div");
        section.className = "cpcard";
        section.style.marginTop = "10px";
        section.style.padding = "10px";

        section.innerHTML = `
        <div class="row"><div><label class="label"> ${i}</label></div></div>

        <label class="label" for="cp_${i}_name">Counterparty name</label>
        <input id="cp_${i}_name" type="text" placeholder="Counterparty name" />

        <label class="label" for="cp_${i}_reason">Reason for sampled</label>
        <textarea id="cp_${i}_reason" type="text" placeholder="Reason for sampled"></textarea>

        <label class="label" for="cp_${i}_addl">Additional information</label>
        <textarea id="cp_${i}_addl"  placeholder="Additional information"></textarea>
      `;
        wrapEl.appendChild(section);
      }
    };

    btnEl.addEventListener("click", build);

    // Optional: press Enter in the number input to trigger Add
    countEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter") build();
    });
  }

  async function copySummary() {
    const html = summaryEditor.innerHTML;
    const text = summaryEditor.innerText;
    if (!text.trim()) return;

    try {
      if (navigator.clipboard && window.ClipboardItem) {
        await navigator.clipboard.write([
          new ClipboardItem({
            "text/html": new Blob([html], { type: "text/html" }),
            "text/plain": new Blob([text], { type: "text/plain" }),
          }),
        ]);
      } else {
        await navigator.clipboard.writeText(text);
      }
      setStatus("Copied.", "ok");
    } catch {
      setStatus("Copy blocked—select and copy manually.", "bad");
    }
  }
  function buildIndividualFocalSummary({ valOrNA, valTrim }) {
    const name = valOrNA("ind_focalName");
    const dobRaw = valTrim("ind_dob");
    const yob =
      dobRaw && !Number.isNaN(new Date(dobRaw).getTime())
        ? String(new Date(dobRaw).getFullYear())
        : "N/A";

    const city = valOrNA("ind_city");
    const state = valOrNA("ind_state");
    const occupation = valOrNA("ind_occupation");
    const employer = valOrNA("ind_employer");
    const relStart = valOrNA("ind_relationshipStart");

    const addlInfo = valTrim("ind_addlInfo");
    const addlInfoClause = addlInfo ? ` ${addlInfo}` : "";

    const customerStatusEl = document.getElementById("ind_customerStatus");
    const endDateRaw = valTrim("ind_endDate");
    const reasonRaw = valTrim("ind_reason");

    const customerStatus =
      (customerStatusEl && customerStatusEl.value) ||
      (endDateRaw || reasonRaw ? "inactive" : "active");

    let statusSentence = "";
    if (customerStatus === "inactive") {
      const endDate = valOrNA("ind_endDate");
      const reason = valOrNA("ind_reason");

      if (endDate === "N/A" || reason === "N/A") {
        alert('Please fill "No longer customer as of" date and reason.');
      }
      statusSentence = `${name} is no longer a TD customer as of ${endDate} because ${reason}.`;
    } else {
      statusSentence = `${name}'s accounts remain active.`;
    }

    return (
      ` ${name}, born in ${yob}, is located in ${city}, ${state} and has been a TD Bank customer since ${relStart}. ` +
      `${name} is a ${occupation} for ${employer}, per internal records${addlInfoClause}. ` +
      `${statusSentence} `
    );
  }
  function buildEntityFocalSummary({ valOrNA, valTrim }) {
    const name = valOrNA("ent_focalName");
    const incorp = valOrNA("ent_incorpDate");
    const city = valOrNA("ent_city");
    const state = valOrNA("ent_state");
    const relStart = valOrNA("ent_relationshipStart");
    const nature = valOrNA("ent_natureBusiness");
    const naics = valOrNA("ent_naics");
    const relatedParty = valOrNA("ent_relatedParty");

    const addlInfo = valTrim("ent_addlInfo");
    const addlInfoClause = addlInfo
      ? ` Additional information: ${addlInfo}.`
      : "";
    const relatedpartypercentage = "input to add";
    const customerStatusEl = document.getElementById("ent_customerStatus");
    const endDateRaw = valTrim("ent_endDate");
    const reasonRaw = valTrim("ent_reason");

    const customerStatus =
      (customerStatusEl && customerStatusEl.value) ||
      (endDateRaw || reasonRaw ? "inactive" : "active");

    let statusSentence = "";
    if (customerStatus === "inactive") {
      const endDate = valOrNA("ent_endDate");
      const reason = valOrNA("ent_reason");

      if (endDate === "N/A" || reason === "N/A") {
        alert('Please fill "No longer customer as of" date and reason.');
      }
      statusSentence = `${name} is no longer a TD customer as of ${endDate} because ${reason}.`;
    } else {
      statusSentence = `${name}'s accounts remain active.`;
    }

    return `${name}, incorporated on ${incorp}, is located in ${city}, ${state} and has been a TD Bank customer since ${relStart}. 
     ${name}'s nature of business is ${nature} (NAICS: ${naics}), per internal records
       ${addlInfoClause}. ${relatedParty} is Listed as the beneficial owner ${relatedpartypercentage} ${statusSentence} `;
  }
  function clearAll() {
    rows = [];
    fileName = "";
    fileEl.value = "";

    summaryEditor.innerHTML = "";

    btnGenerate.disabled = true;
    btnCopy.disabled = true;
    btnClear.disabled = true;

    setStatus("", "");
  }

  const valOrNA = (id) =>
    (document.getElementById(id)?.value || "").trim() || "N/A";
  const valTrim = (id) => (document.getElementById(id)?.value || "").trim();
  // Main Summary
  function buildSummaryHtml() {
    if (!rows.length) throw new Error("No transactions loaded.");

    const focalType =
      document.getElementById("focalType")?.value || "individual";

    let focalentitysummary = "";
    if (focalType === "individual") {
      focalentitysummary = buildIndividualFocalSummary({
        valOrNA,
        valTrim,
      });
    } else if (focalType === "entity") {
      focalentitysummary = buildEntityFocalSummary({
        valOrNA,
        valTrim,
      });
    } else {
      alert("Please check all the input fields are selected correctly.");
      focalentitysummary = "";
    }

    // If you have different name ids by type, pick accordingly:
    const name =
      focalType === "entity"
        ? valOrNA("ent_focalName")
        : valOrNA("ind_focalName");

    const Rfisummary = buildRfiSummary(name);
    const counterpartyLines = buildCounterpartySummary();
    const analystsummary = buildAnalystDisposition();
    const summaryConclusion = Conclusion();
    const ptopsummarydisposition = p2psummary();
    const cashSummarydisposition = cashSummary();
    return `
    <div>
    <Strong>Tranche Analysis</Strong>
    <p>${esc(focalentitysummary)}</p>
    <Strong>RFI</Strong>
    <p>${esc(Rfisummary)}</p>
    <Strong>Counterparty additional info:</Strong>
    ${counterpartyLines.map((l) => `<p>${esc(l)}</p>`).join("")}
    <Strong>Case Decision</Strong>
    <p>${analystsummary}</p>
    <Strong>P2P Transaction Analysis</Strong>
    <p>${ptopsummarydisposition}</p>
    <Strong>Cash Transaction Analysis</Strong>
    <p>${cashSummarydisposition}</p>
    <Strong>Conclusion</Strong> 
      <p>${summaryConclusion}</p>
     
    </div>
  `.trim();
  }
  clearAll();
  wireCounterpartyAddButton();
  wireCustomerStatusToggle();
  wireRfiStatusFields();
  setFocalFormVisibility();
  document.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-cmd], [data-action]");
    if (!btn) return;

    const targetId = btn.getAttribute("data-target");
    const target = targetId ? document.getElementById(targetId) : null;
    if (!target) return;

    target.focus();

    const cmd = btn.getAttribute("data-cmd");
    const action = btn.getAttribute("data-action");

    if (cmd) document.execCommand(cmd, false, null);
    if (action === "clearFormat") {
      document.execCommand("removeFormat", false, null);
      document.execCommand("unlink", false, null);
    }
  });

  summaryEditor.addEventListener("input", () => {
    btnCopy.disabled = !summaryEditor.textContent.trim();
  });

  fileEl.addEventListener("change", async () => {
    setStatus("", "");
    const f = fileEl.files?.[0];
    if (!f) return;

    try {
      if (!window.XLSX)
        throw new Error("Missing XLSX library (xlsx.full.min.js).");

      fileName = f.name;
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });

      const sheetName = findSheetByName(wb, "Transactions");
      if (!sheetName) {
        const available = (wb.SheetNames || []).join(", ") || "(none)";
        throw new Error(
          `Sheet "Transactions" not found. Available: ${available}`,
        );
      }

      const ws = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!json.length)
        throw new Error('No rows found in "Transactions" sheet.');

      rows = json;
      TransactionDataPoints(rows);
      btnGenerate.disabled = false;
      btnClear.disabled = false;

      setStatus(`Loaded: ${f.name} (Sheet: ${sheetName})`, "ok");
    } catch (err) {
      clearAll();
      setStatus(err?.message || "Failed to load file.", "bad");
    }
  });

  btnGenerate.addEventListener("click", () => {
    try {
      const html = buildSummaryHtml();
      summaryEditor.innerHTML = html;
      btnCopy.disabled = !summaryEditor.textContent.trim();
      setStatus("Generated. Edit the summary as needed.", "ok");
    } catch (err) {
      setStatus(err?.message || "Failed to generate summary.", "bad");
    }
  });
  focalTypeEl?.addEventListener("change", setFocalFormVisibility);

  btnCopy.addEventListener("click", copySummary);
  btnClear.addEventListener("click", clearAll);
})();
