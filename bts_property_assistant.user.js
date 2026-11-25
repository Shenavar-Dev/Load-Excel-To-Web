// ==UserScript==
// @name         BTS Property Excel Assistant Deep
// @namespace    https://btsproperty.mtnirancell.ir/
// @author       Ebrahim Shenavar | Irancell
// @version      6.5.6
// @description  Inline panel (max-width 1020px) under topmenu-nav; safe Excel review and form population (v6.5.6)
// @match        https://btsproperty.mtnirancell.ir/Site_Action.aspx*
// @downloadURL  https://github.com/Shenavar-Dev/Load-Excel-To-Web/raw/refs/heads/main/bts_property_assistant.user.js
// @updateURL    https://github.com/Shenavar-Dev/Load-Excel-To-Web/raw/refs/heads/main/bts_property_assistant.user.js
// @grant        none
// @run-at       document-idle
// ==/UserScript==

/*
 * Version 6.5.4:
 * - اصلاح منطق جستجوی Site ID: معیار تأیید وجود سایت، عدم نمایش باکس ثبت سایت جدید (Content_Site / Content_Txt_Site_ID_insert) به‌جای وابستگی به پر شدن LL Name قرار گرفت.
 *
 * Version 6.5.5:
 * - تصفیه خطاهای فرمولی اکسل (#N/A، #VALUE!، #REF!، #DIV/0!، #NAME?، #NULL!، #NUM!، #GETTING_DATA) و برخورد مشابه خانه خالی تا دادهٔ صحیح وب بازنویسی نشود.
 * - افزودن چک‌باکس انتخاب/غیرفعال‌سازی فیلدها در تنظیمات نگاشت هدر (به‌جز فیلدهای الزامی هویتی) و ذخیره وضعیت در localStorage.
 *
 * Version 6.5.6:
 * - تفکیک پیام وضعیت بین «خالی عمدی» و «خطای فرمول اکسل» و نمایش مقدار «فعلی در وب» زیر باکس در حالت‌های خالی، خطای فرمول، نامعتبر و غیرقابل ویرایش.
 * - افزایش یک‌پله‌ای اندازه فونت‌ها و بهبود سلسله‌مراتب خوانایی پنل بدون به‌هم‌ریختن چیدمان چندستونی در عرض 1020px.
 * - بازبینی خودکار هویت هنگام باز شدن مجدد پنل و مسدودسازی کامل‌تر Enter برای جلوگیری از PostBack ناخواسته ASP.NET.
 */


(function () {
    'use strict';

    if (document.getElementById('epf-root')) return;

    // ============================================================
    // 01. FIELD DEFINITIONS & HEADER MAPPING
    // ============================================================
    const WEB_FIELDS = [
        { key: 'siteId', label: 'Site ID', id: 'Content_Txt_SiteID', type: 'input', group: 'identity' },
        { key: 'contractType', label: 'Contract Type', id: 'ddl_Contract_type', type: 'select', group: 'identity' },
        { key: 'llName', label: 'LL Name', id: 'Content_txt_LLNAME', type: 'input', group: 'identity' },
        { key: 'ownerType', label: 'Owner Type', id: 'Content_DDL_OWNER_TYPE', type: 'select', group: 'identity' },
        { key: 'startDate', label: 'Start Date', id: 'Content_TB_LAST_LEASE_START_SHAMSI', type: 'input', group: 'dates' },
        { key: 'endDate', label: 'End Date', id: 'Content_TB_LAST_LEASE_END_SHAMSI', type: 'input', group: 'dates' },
        { key: 'taxType', label: 'Tax Type', id: 'Content_DDL_TAX_TYPE', type: 'select', group: 'finance' },
        { key: 'contractAmount', label: 'Contract Gross Yearly Base', id: 'Content_TB_CONTRACT_AMOUNT', type: 'input', group: 'finance' },
        { key: 'escalationType', label: 'Escalation Type', id: 'Content_DDL_ESCALATION_Type', type: 'select', group: 'finance' },
        { key: 'escalation', label: 'Escalation %', id: 'Content_DDL_ESCALATION', type: 'select', group: 'finance' },
        { key: 'llCount', label: 'LL Count', id: 'Content_DDL_LL_COUNT', type: 'select', group: 'finance' },
        { key: 'paymentTerm', label: 'Payment Term', id: 'Content_DDL_PAYMENT_TERM', type: 'select', group: 'payment' },
        { key: 'advancePayment', label: 'Advance Payment', id: 'Content_TB_ADVANCE_PAYMENT', type: 'input', group: 'payment' },
        { key: 'paymentComment', label: 'Payment Comment', id: 'Content_TB_PAYMENT_COMMENT', type: 'input', group: 'payment' },
        { key: 'paymentStatus', label: 'Payment Status', id: 'Content_DDL_PAYMENT_STATUS', type: 'select', group: 'payment' },
        { key: 'prtRegion', label: 'PRT Region', id: 'Content_DDL_PRT_REGION', type: 'select', group: 'region' },
        { key: 'mpZone', label: 'MP Zone', id: 'Content_TB_MP_ZONE', type: 'input', group: 'region' },
        { key: 'installType', label: 'Install Type', id: 'Content_DDL_INSTALL_TYPE', type: 'select', group: 'region' },
        { key: 'depends', label: 'Depends', id: 'Content_TB_DEPENDS', type: 'input', group: 'region' },
        { key: 'pType', label: 'P Type', id: 'Content_DDL_P_TYPE', type: 'select', group: 'region' },
        { key: 'nationalId', label: 'National ID', id: 'Content_TB_NATIONAL_ID', type: 'input', group: 'owner' },
        { key: 'llManner', label: 'LL Manner', id: 'Content_DDL_LL_MANNER', type: 'select', group: 'owner' },
        { key: 'legalHistory', label: 'Legal History', id: 'Content_DDL_LEGAL_HISTORY', type: 'select', group: 'owner' },
        { key: 'accessTypeLease', label: 'Access Type Lease', id: 'Content_DDL_ACCESS_TYPE_LEASE_BASE', type: 'select', group: 'owner' },
        { key: 'accessTypeLL', label: 'Access Type LL', id: 'Content_DDL_ACCESS_TYPE_LL_BASE', type: 'select', group: 'owner' },
        { key: 'address', label: 'Address', id: 'Content_TB_ADDRESS', type: 'textarea', group: 'owner' },
        { key: 'phoneNumber', label: 'Phone Number', id: 'Content_TB_PHONE_NUMBER', type: 'input', group: 'owner' },
        { key: 'llChanges', label: 'LL Changes', id: 'Content_DDL_LL_CHANGES', type: 'select', group: 'status' },
        { key: 'signatureStatus', label: 'Signature Status', id: 'Content_DDL_SIGNATURE_STATUS', type: 'select', group: 'status' },
        { key: 'renewalStatus', label: 'Renewal Status', id: 'Content_DDL_RENEWAL_STATUS', type: 'select', group: 'status' },
        { key: 'leaseType', label: 'Lease Type', id: 'Content_DDL_LEASE_TYPE', type: 'select', group: 'status' },
        { key: 'renewalOwner', label: 'Renewal Owner', id: 'Content_TB_RENEWAL_OWNER', type: 'input', group: 'status' },
        { key: 'cancellationLetter', label: 'Cancellation Letter', id: 'Content_DDL_CANCELLATION_LETTER', type: 'select', group: 'status' },
        { key: 'cancellationDate', label: 'Cancellation Letter Submission Date', id: 'Content_TB_CANCELLATION_LETTER_SUBMISSION_DATE', type: 'input', group: 'status' },
        { key: 'cancellationComment', label: 'Cancellation Letter Comment', id: 'Content_TB_CANCELLATION_LETTER_COMMENT', type: 'textarea', group: 'status' },
        { key: 'extraComment', label: 'Extra Comment', id: 'Content_TB_EXTRA_COMMENT', type: 'textarea', group: 'status' }
    ];

    const FIELD_BY_KEY = Object.create(null);
    WEB_FIELDS.forEach(field => { FIELD_BY_KEY[field.key] = field; });

    const FIELD_KEYS = {
        SITE_ID: 'siteId', CONTRACT_TYPE: 'contractType', AMOUNT: 'contractAmount',
        ESCALATION_TYPE: 'escalationType', START_DATE: 'startDate', END_DATE: 'endDate'
    };
    const CALC_FIELD_KEYS = ['taxType', 'contractAmount', 'escalationType', 'escalation', 'llCount'];
    const HEADER_STORAGE_KEY = 'epf:bts-excel-paste:header-row:v1';
    const DISABLED_FIELDS_STORAGE_KEY = 'epf:bts-excel-paste:disabled-fields:v1';

    // Regex to match common Excel formula errors
    const EXCEL_ERROR_PATTERN = /^#(N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NULL!|NUM!|GETTING_DATA)$/i;

    function parseExcelCell(val) {
        if (val == null) return { value: '', isError: false, rawError: '' };
        const str = String(val).trim();
        if (EXCEL_ERROR_PATTERN.test(str)) {
            return { value: '', isError: true, rawError: str };
        }
        return { value: str, isError: false, rawError: '' };
    }

    function loadStoredDisabledFields() {
        try {
            const json = window.localStorage.getItem(DISABLED_FIELDS_STORAGE_KEY);
            if (!json) return new Set();
            const arr = JSON.parse(json);
            return new Set(Array.isArray(arr) ? arr : []);
        } catch (e) {
            return new Set();
        }
    }

    function persistDisabledFields(disabledSet) {
        try {
            const arr = Array.from(disabledSet);
            window.localStorage.setItem(DISABLED_FIELDS_STORAGE_KEY, JSON.stringify(arr));
            return true;
        } catch (e) {
            return false;
        }
    }

    let activeDisabledFields = loadStoredDisabledFields();

    function isFieldEnabled(key) {
        if (key === FIELD_KEYS.SITE_ID || key === FIELD_KEYS.CONTRACT_TYPE) return true;
        return !activeDisabledFields.has(key);
    }

    // Normalises a header cell for case- and space-insensitive matching.
    function normalizeHeader(value) {
        return String(value == null ? '' : value).replace(/^\uFEFF/, '').trim().replace(/\s+/g, ' ').toLowerCase();
    }

    // Converts Latin digits to Persian digits for display.
    function faNum(value) {
        return String(value).replace(/\d/g, digit => '۰۱۲۳۴۵۶۷۸۹'[digit]);
    }

    const REQUIRED_HEADER_KEYS = [FIELD_KEYS.SITE_ID, FIELD_KEYS.CONTRACT_TYPE];

    // Maps available fields to their Excel columns and validates minimum required headers.
    function buildHeaderMapping(headerText) {
        const parsedRows = parseTSV(headerText);
        if (!parsedRows.length || parsedRows.length > 1) {
            return { parsed: false, ready: false, missing: 0, mappedCount: 0 };
        }

        const normalized = parsedRows[0].map(normalizeHeader);

        const columns = Object.create(null);
        const states = Object.create(null);
        let missing = 0;
        let dupCount = 0;
        let mappedCount = 0;
        let enabledCount = 0;

        WEB_FIELDS.forEach(field => {
            if (!isFieldEnabled(field.key)) {
                states[field.key] = { st: 'disabled', col: null };
                return;
            }
            enabledCount += 1;
            const wanted = normalizeHeader(field.label);
            const matches = [];
            normalized.forEach((name, index) => {
                if (name === wanted) matches.push(index);
            });
            if (!matches.length) {
                states[field.key] = { st: 'missing', col: null };
                missing += 1;
                return;
            }
            columns[field.key] = matches[0];
            mappedCount += 1;
            if (matches.length > 1) dupCount += 1;
            states[field.key] = { st: matches.length > 1 ? 'dup' : 'ok', col: matches[0] };
        });

        const hasSiteId = Number.isInteger(columns[FIELD_KEYS.SITE_ID]);
        const hasContractType = Number.isInteger(columns[FIELD_KEYS.CONTRACT_TYPE]);
        const ready = hasSiteId && hasContractType && mappedCount >= 3;

        return { parsed: true, ready, missing, mappedCount, enabledCount, hasSiteId, hasContractType, columns, states, dupCount };
    }

    // Identity mapping: default column order equals the field index.
    function defaultMapping() {
        const columns = Object.create(null);
        const states = Object.create(null);
        let enabledCount = 0;
        WEB_FIELDS.forEach((field, index) => {
            if (!isFieldEnabled(field.key)) {
                states[field.key] = { st: 'disabled', col: null };
                return;
            }
            enabledCount += 1;
            columns[field.key] = index;
            states[field.key] = { st: 'ok', col: index };
        });
        return { parsed: true, ready: true, missing: 0, mappedCount: enabledCount, enabledCount, columns, states, dupCount: 0 };
    }

    function isDefaultMapping(mapping) {
        return Boolean(mapping) && WEB_FIELDS.every((field, index) => {
            if (!isFieldEnabled(field.key)) return true;
            return mapping.columns[field.key] === index;
        });
    }

    function loadStoredHeader() {
        try {
            const value = window.localStorage.getItem(HEADER_STORAGE_KEY);
            return value == null ? '' : value;
        } catch (error) {
            initNote = 'ذخیره‌سازی مرورگر در دسترس نیست؛ هدر فقط در همین نشست قابل استفاده است.';
            return '';
        }
    }

    function persistHeader(headerText) {
        try {
            window.localStorage.setItem(HEADER_STORAGE_KEY, headerText);
            return true;
        } catch (error) {
            return false;
        }
    }

    // Non-blocking extraction; cleans Excel formula errors and tracks out-of-range mapped columns.
    function extractValues(cells, mapping) {
        if (!mapping) {
            return { ok: false, message: 'ابتدا هدر اکسل را در «تنظیمات نگاشت هدر» ثبت کنید.' };
        }

        const webValues = {};
        const excelErrors = Object.create(null);
        const outOfRange = Object.create(null);
        WEB_FIELDS.forEach(field => {
            if (!isFieldEnabled(field.key)) return;
            const col = mapping.columns[field.key];
            if (!Number.isInteger(col)) return;
            if (col >= cells.length) {
                outOfRange[field.key] = true;
                return;
            }
            const parsedCell = parseExcelCell(cells[col]);
            if (parsedCell.isError) {
                excelErrors[field.key] = parsedCell.rawError;
                webValues[field.key] = '';
                return;
            }
            webValues[field.key] = field.type === 'textarea' ? parsedCell.value : parsedCell.value.replace(/\n/g, ' ');
        });
        return { ok: true, webValues, excelErrors, outOfRange };
    }

    function parseTSV(text) {
        const source = String(text || '').replace(/\r\n?/g, '\n').replace(/[\u2028\u2029]/g, '\n');
        const rows = [];
        let row = [];
        let cell = '';
        let quoted = false;
        let cellStart = true;

        for (let index = 0; index < source.length; index++) {
            const char = source[index];
            if (quoted) {
                if (char === '"') {
                    if (source[index + 1] === '"') { cell += '"'; index++; }
                    else quoted = false;
                } else cell += char;
                continue;
            }
            if (char === '"' && cellStart) { quoted = true; cellStart = false; continue; }
            if (char === '\t') { row.push(cell); cell = ''; cellStart = true; continue; }
            if (char === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; cellStart = true; continue; }
            cell += char;
            cellStart = false;
        }
        row.push(cell);
        rows.push(row);
        return rows.filter(item => !(item.length === 1 && item[0].trim() === ''));
    }

    const GROUPS = [
        { id: 'identity', title: 'هویت قرارداد' },
        { id: 'dates', title: 'تاریخ‌ها' },
        { id: 'finance', title: 'مالی و محاسبه' },
        { id: 'payment', title: 'پرداخت' },
        { id: 'region', title: 'منطقه و نصب' },
        { id: 'owner', title: 'مالک و تماس' },
        { id: 'status', title: 'وضعیت و توضیحات' }
    ];

    const MODAL_ROWS = [
        { id: 'identity', rows: [{ fields: ['siteId', 'contractType'], facing: true }, { fields: ['llName', 'ownerType'], facing: true }] },
        { id: 'dates', rows: [{ fields: ['startDate', 'endDate'], facing: true }] },
        { id: 'finance', rows: [{ fields: ['taxType', 'contractAmount', 'escalationType'] }, { fields: ['escalation', 'llCount'], facing: true }] },
        { id: 'payment', rows: [{ fields: ['paymentTerm', 'advancePayment', 'paymentStatus'] }, { fields: ['paymentComment'] }] },
        { id: 'region', rows: [{ fields: ['prtRegion', 'mpZone', 'installType'] }, { fields: ['depends', 'pType'], facing: true }] },
        { id: 'owner', rows: [{ fields: ['nationalId', 'llManner', 'phoneNumber'] }, { fields: ['legalHistory', 'accessTypeLease', 'accessTypeLL'] }, { fields: ['address'] }] },
        { id: 'status', rows: [{ fields: ['llChanges', 'signatureStatus', 'renewalStatus'] }, { fields: ['leaseType', 'renewalOwner', 'cancellationLetter'] }, { fields: ['cancellationDate', 'cancellationComment'] }, { fields: ['extraComment'] }] }
    ];

    const CALC_BTN_ID = 'Content_btn_CalculateAmount';
    const LL_NAME_ID = 'Content_txt_LLNAME';
    const SEARCH_BTN_ID = 'Content_btn_Search';
    const SITE_ID_INPUT_ID = 'Content_Txt_SiteID';
    const SITE_INSERT_INPUT_ID = 'Content_Txt_Site_ID_insert';
    const DATE_OPTIONAL_TYPES = ['NOT DELIVERED', 'SITE SHARING', 'PURCHASED'];

    let initNote = '';
    let activeHeaderText = loadStoredHeader();
    let activeHeaderMapping = null;
    if (activeHeaderText) {
        const bootResult = buildHeaderMapping(activeHeaderText);
        if (bootResult.parsed && bootResult.ready) {
            activeHeaderMapping = { columns: bootResult.columns };
        } else initNote = 'هدر ذخیره‌شده معتبر نیست؛ از تنظیمات نگاشت هدر دوباره ثبت کنید.';
    }

    const state = {
        webValues: {},
        excelErrors: {},
        outOfRange: {},
        parsed: false,
        reviewedText: '',
        searchConfirmed: false,
        headerContract: '',
        contractChangeCompleted: false,
        confirmOpen: false,
        isOpen: false,
        settingsOpen: false,
        refreshQueued: false,
        searching: false,
        calculating: false,
        cancelRequested: false
    };

    function getVal(key) {
        const value = state.webValues[key];
        return value === undefined ? '' : value;
    }

    const controls = Object.create(null);
    const rows = Object.create(null);
    const groupEls = [];
    const modalRowEls = [];
    let built = false;
    let contractFixTimer = null;
    let ctBox = null;
    let ctBtn = null;
    let ctMsg = null;
    let headerBtn = null;

    // ============================================================
    // 02. TEXT & DOM UTILITIES
    // ============================================================

    function fixChars(s) {
        return s.replace(/\u064A/g, '\u06CC').replace(/\u0643/g, '\u06A9');
    }

    function normalize(value) {
        return fixChars(String(value == null ? '' : value)).trim().toLowerCase();
    }

    function normalizeStrong(value) {
        return fixChars(String(value == null ? '' : value))
            .replace(/\u200c/g, '')
            .replace(/\s+/g, '')
            .trim()
            .toLowerCase();
    }

    function looseText(value) {
        return fixChars(String(value == null ? '' : value))
            .replace(/[\u200e\u200f]/g, '')
            .replace(/\s+/g, ' ')
            .trim()
            .toLowerCase();
    }

    function stripCommas(value) {
        return String(value == null ? '' : value).replace(/,/g, '').trim();
    }

    function makeElement(tagName, className, text) {
        const el = document.createElement(tagName);
        if (className) el.className = className;
        if (text !== undefined) el.textContent = text;
        return el;
    }

    function isElementDisabled(el) {
        if (!el) return true;
        return Boolean(el.disabled || el.readOnly ||
            (typeof el.matches === 'function' && el.matches(':disabled')));
    }

    function getPRM() {
        try {
            const w = window.Sys && window.Sys.WebForms;
            return w && w.PageRequestManager ? w.PageRequestManager.getInstance() : null;
        } catch (e) {
            return null;
        }
    }

    function targetInfo(field) {
        const el = field.id ? document.getElementById(field.id) : null;
        if (!el) return { element: null, exists: false, typeMatches: false, editable: false };

        const tag = el.tagName.toLowerCase();
        const typeMatches = field.type === 'input' ? tag === 'input'
            : field.type === 'select' ? tag === 'select'
            : field.type === 'textarea' ? tag === 'textarea'
            : false;

        if (!typeMatches) return { element: el, exists: true, typeMatches: false, editable: false };
        return { element: el, exists: true, typeMatches: true, editable: !isElementDisabled(el) };
    }

    function findSelectableOption(selectEl, value) {
        if (!selectEl || selectEl.tagName.toLowerCase() !== 'select') return null;
        const wanted = normalize(value);
        const wantedStrong = normalizeStrong(value);

        const options = Array.from(selectEl.options || []).filter(opt => {
            const group = opt.parentElement;
            return !opt.disabled && !(group && group.tagName && group.tagName.toLowerCase() === 'optgroup' && group.disabled);
        });

        let match = options.find(opt => normalize(opt.value) === wanted);
        if (match) return match;

        match = options.find(opt => normalize(opt.textContent) === wanted);
        if (match) return match;

        match = options.find(opt => normalizeStrong(opt.value) === wantedStrong);
        if (match) return match;

        match = options.find(opt => normalizeStrong(opt.textContent) === wantedStrong);
        return match || null;
    }

    function findOptionByText(selectEl, value) {
        if (!selectEl || selectEl.tagName.toLowerCase() !== 'select') return null;
        const options = Array.from(selectEl.options || []).filter(o => !o.disabled);
        const w = normalize(value);
        let m = options.find(o => normalize(o.textContent) === w);
        if (m) return m;
        const ws = normalizeStrong(value);
        return options.find(o => normalizeStrong(o.textContent) === ws) || null;
    }

    function selectedOptionText(selectEl) {
        if (!selectEl || selectEl.tagName.toLowerCase() !== 'select') return '';
        const selected = selectEl.selectedOptions && selectEl.selectedOptions.length
            ? selectEl.selectedOptions[0]
            : (selectEl.selectedIndex >= 0 ? selectEl.options[selectEl.selectedIndex] : null);
        return selected ? String(selected.textContent || '').trim() : '';
    }

    function getCurrentFormValue(field) {
        const info = targetInfo(field);
        if (!info.exists || !info.typeMatches) return null;
        if (field.type === 'select') {
            return selectedOptionText(info.element);
        }
        return String(info.element.value || '').trim();
    }

    function resolveCalcAmount() {
        const rowAmount = stripCommas(getVal(FIELD_KEYS.AMOUNT));
        if (rowAmount) {
            const n = Number(rowAmount);
            return { source: 'row', valid: !isNaN(n) && n > 0 };
        }
        const info = targetInfo(FIELD_BY_KEY[FIELD_KEYS.AMOUNT]);
        const formAmount = info.exists ? stripCommas(info.element.value) : '';
        const n = Number(formAmount);
        return { source: 'form', valid: formAmount !== '' && !isNaN(n) && n > 0 };
    }

    function fieldSat(field, pastedValue) {
        const info = targetInfo(field);
        if (!info.exists || !info.typeMatches) return false;
        if (field.type === 'select') {
            const opt = findSelectableOption(info.element, pastedValue);
            if (!opt) return false;
            return String(info.element.value) === String(opt.value);
        }
        if (field.key === FIELD_KEYS.AMOUNT) {
            return stripCommas(info.element.value) === stripCommas(pastedValue);
        }
        return looseText(info.element.value) === looseText(pastedValue);
    }

    // ============================================================
    // 03. ASP.NET REQUEST HANDLING
    // ============================================================

    let activeWait = null;

    function cancelActiveWait() {
        state.cancelRequested = true;
        if (activeWait) activeWait.cancel();
        state.searching = false;
        state.calculating = false;
        setBusy(false);
    }

    function runAndWait(action, startMs, maxMs, onSlow, options) {
        const opts = options || {};
        return new Promise(resolve => {
            const prm = getPRM();
            if (!prm) {
                let timer = null;
                let finished = false;
                const finishLocal = result => {
                    if (finished) return;
                    finished = true;
                    clearTimeout(timer);
                    if (activeWait === handle) activeWait = null;
                    resolve(result);
                };
                const handle = { cancel: () => finishLocal({ started: null, error: null, cancelled: true }) };
                activeWait = handle;
                try { action(); } catch (error) { finishLocal({ started: false, error }); return; }
                timer = setTimeout(() => finishLocal({ started: null, error: null }), 2500);
                return;
            }

            let started = false;
            let done = false;
            let startTimer;
            let slowTimer;
            let maxTimer;

            const finish = result => {
                if (done) return;
                done = true;
                clearTimeout(startTimer);
                clearTimeout(slowTimer);
                clearTimeout(maxTimer);
                if (activeWait === handle) activeWait = null;
                try { prm.remove_beginRequest(onBegin); } catch (error) { }
                try { prm.remove_endRequest(onEnd); } catch (error) { }
                setTimeout(() => resolve(result), 150);
            };
            const handle = {
                cancel: () => {
                    if (opts.abortOnCancel) {
                        try { if (prm.get_isInAsyncPostBack()) prm.abortPostBack(); } catch (error) { }
                    }
                    finish({ started, error: null, cancelled: true });
                }
            };
            const onBegin = () => {
                started = true;
                clearTimeout(startTimer);
            };
            const onEnd = (sender, args) => {
                if (!started) return;
                let error = null;
                try { error = args && args.get_error ? args.get_error() : null; } catch (readError) { error = readError; }
                finish({ started: true, error });
            };

            activeWait = handle;
            prm.add_beginRequest(onBegin);
            prm.add_endRequest(onEnd);
            startTimer = setTimeout(() => {
                if (!started) finish({ started: false, error: null });
            }, startMs);
            slowTimer = setTimeout(() => {
                if (!done && onSlow) onSlow();
            }, 4000);
            maxTimer = setTimeout(() => finish({ started, error: null, timeout: true }), maxMs);

            try { action(); } catch (error) { finish({ started: false, error }); }
        });
    }

    // ============================================================
    // 04. STYLES (Upgraded Typography & Readability)
    // ============================================================

    const styleElement = makeElement('style');
    styleElement.id = 'epf-style';
    styleElement.textContent = `
        #epf-root, #epf-root * { box-sizing: border-box; }
        #epf-root {
            --pri: #004E5B; --pri-d: #003D47; --acc: #F4C300; --ink: #243238; --mut: #526167;
            --line: #dbe2e5; --ok: #218838; --warn: #9a7b0a; --bad: #c82333;
            --header-height: 64px; --topmenu-height: 56px;
            font-family: Irancell, Tahoma, sans-serif; direction: rtl;
        }
        #epf-root [hidden] { display: none !important; }
        #epf-root button, #epf-root input, #epf-root select, #epf-root textarea { font-family: inherit; }
        .swal2-container { z-index: 20000 !important; }

        /* Header button injected as a <li> inside #topmenuNav */
        #epfMenuLink { list-style: none; display: inline-flex; align-items: center; }
        #epf-header-btn {
            display: inline-flex; align-items: center; justify-content: center; gap: 6px;
            padding: 7px 14px; border: 1px solid #b8dce8; border-radius: 8px;
            background: #e6f3f8; color: var(--pri);
            font-size: 13.5px; font-weight: 700; cursor: pointer; white-space: nowrap;
            transition: background .15s ease, box-shadow .15s ease;
            direction: rtl; text-decoration: none;
        }
        #epf-header-btn:hover { background: #d4ebf4; color: var(--pri); }
        #epf-header-btn:focus-visible { outline: 2px solid var(--pri); outline-offset: 2px; }
        #epf-header-btn[aria-expanded="true"] { background: #c6e8f4; font-weight: 700; }
        #epf-header-btn .link-text { font-size: 13.5px; font-weight: 700; }
        #epf-header-btn i { font-size: 14px; line-height: 1; }

        /* Inline panel (not floating) — max-width 1020px */
        #epf-overlay {
            display: none;
            width: 100%;
            max-width: 1020px;
            margin: 8px auto 14px;
            overflow: hidden;
            direction: rtl;
            opacity: 0;
            transition: opacity .22s ease;
            scroll-margin-top: calc(var(--header-height, 64px) + var(--topmenu-height, 56px) + 20px);
        }
        #epf-overlay.epf-open {
            display: block;
            opacity: 1;
        }
        #epf-overlay.epf-closing {
            opacity: 0;
        }
        #epf-modal {
            width: 100%; display: flex; flex-direction: column;
            background: #fff; color: var(--ink); border-radius: 12px;
            border: 1px solid var(--line); box-shadow: 0 4px 14px rgba(0,0,0,.08);
            direction: rtl; overflow: hidden;
        }
        .epf-header { position: relative; flex: 0 0 auto; padding: 11px 16px 10px; background: var(--pri); color: #fff; }
        .epf-header-top { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
        .epf-title-wrap { display: flex; align-items: center; gap: 8px; min-width: 0; }
        .epf-title { margin: 0; font-size: 16.5px; font-weight: 700; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .epf-identity { flex: 0 1 auto; max-width: 340px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
            direction: ltr; font-size: 13px; font-weight: 700; color: #a5d6a7; background: rgba(255,255,255,.12); padding: 3px 11px; border-radius: 12px; }
        .epf-xbtn { flex: 0 0 auto; width: 28px; height: 28px; display: inline-flex; align-items: center; justify-content: center;
            border: 0; border-radius: 6px; background: rgba(255,255,255,.16); color: #fff; font-size: 15px; line-height: 1; cursor: pointer; }
        .epf-xbtn:hover { background: rgba(255,255,255,.3); }
        .epf-subtitle { margin: 5px 0 7px; font-size: 12px; line-height: 1.65; color: rgba(255,255,255,.92); text-align: right; }
        .epf-gear {
            flex: 0 0 auto; display: inline-flex; flex-direction: column; align-items: center; gap: 1px; padding: 5px 11px;
            border: 1px solid #9bd5ff; border-radius: 7px; background: #cceaff; color: #14527f;
            font-size: 12px; font-weight: 700; cursor: pointer;
        }
        .epf-gear:hover:not(:disabled) { background: #b5ddfa; }
        .epf-gear:disabled { opacity: .45; cursor: not-allowed; }
        .epf-gear-main { font-size: 12px; font-weight: 700; }
        .epf-gear-state { font-size: 11px; font-weight: 400; line-height: 1.25; }
        .epf-gear[data-mode="custom"] .epf-gear-state { color: #8a3b00; font-weight: 700; }
        .epf-gear[data-mode="none"] .epf-gear-state { color: #a33; }
        .epf-paste-label { display: block; margin: 0 0 4px; font-size: 12.5px; font-weight: 700; text-align: right; }
        .epf-paste-layout { display: grid; grid-template-columns: minmax(0, 1fr) 150px; gap: 8px; direction: ltr; align-items: stretch; }
        #epf-paste-area {
            display: block; width: 100%; min-height: 50px; max-height: 80px; resize: vertical; padding: 7px 10px;
            border: 1px solid #cbd3d6; border-radius: 7px; background: #fbfcfc; color: var(--ink);
            font-size: 12.5px; line-height: 1.5; direction: ltr; text-align: left; white-space: pre; overflow-wrap: normal;
        }
        .epf-paste-actions { display: flex; flex-direction: row; gap: 6px; }
        .epf-paste-actions .epf-btn { flex: 1; min-width: 0; min-height: 50px; padding: 6px 7px; }
        #epf-map-note, #epf-paste-warning {
            display: none; margin: 0 0 6px; padding: 5px 9px; border-radius: 6px; font-size: 11.5px; line-height: 1.6;
            white-space: pre-wrap;
        }
        #epf-map-note.epf-show { display: block; color: #7a5b00; background: #fff8e6; border: 1px solid #f0d78c; }
        #epf-paste-warning.epf-show { display: block; margin: 6px 0 0; color: #7a5b00; background: #fff8e6; border: 1px solid #f0d78c; }
        .epf-btn { border: 0; border-radius: 7px; cursor: pointer; padding: 8px 15px; font-size: 12.5px; font-weight: 700; }
        .epf-btn:disabled { opacity: .5; cursor: not-allowed; }
        .epf-btn-primary { color: var(--pri); background: var(--acc); }
        .epf-btn-apply { color: #fff; background: var(--ok); }
        .epf-btn-secondary { color: #45545a; background: #e9edef; }
        .epf-btn-cancel { color: #842029; background: #f8d7da; border: 1px solid #f1aeb5; }
        .epf-btn-cancel:hover:not(:disabled) { background: #f1c2c7; }
        .epf-btn-goto-save { color: #fff; background: var(--pri); border: 1px solid var(--pri-d); }
        .epf-btn-goto-save:hover:not(:disabled) { background: var(--pri-d); }
        .epf-btn-reset { width: auto; padding: 6px 12px; font-size: 12px; color: var(--pri); background: #e6f3f8; border: 1px solid #b8dce8; }
        .epf-btn-reset:hover:not(:disabled) { background: #d4ebf4; }
        .epf-content { flex: 1 1 auto; min-height: 0; padding: 11px 15px; overflow: visible; }
        #epf-main-message { display: none; margin: 0 0 9px; padding: 7px 12px; border-radius: 7px; font-size: 12px; line-height: 1.7; white-space: pre-wrap; text-align: center; }
        #epf-main-message.epf-show { display: block; }
        .epf-msg-error { color: #842029; background: #f8d7da; border: 1px solid #f5c6cb; }
        .epf-msg-success { color: #155724; background: #d4edda; }
        .epf-msg-info { color: #174b58; background: #e8f1f3; }
        .epf-msg-warning { color: #7a5b00; background: #fff8e6; border: 1px solid #f0d78c; }
        #epf-fields { display: block; }
        .epf-section { margin-top: 10px; padding: 0 10px 10px; overflow: hidden; border: 1px solid #cce3eb; border-radius: 8px; background: #fdfefe; box-shadow: 0 1px 3px rgba(0,0,0,0.02); }
        .epf-section:first-child { margin-top: 0; }
        .epf-section-title { margin: 0 -10px 8px; padding: 5px 10px; text-align: center; font-size: 13px; font-weight: 700; color: var(--pri);
            background: #e6f3f8; border-bottom: 1px solid #c3dfe9; }
        .epf-mrow { display: flex; direction: ltr; gap: 8px; }
        .epf-mrow + .epf-mrow { margin-top: 7px; }
        .epf-fcell { --w: 1 1 0; flex: var(--w); min-width: 0; display: flex; direction: ltr; gap: 7px; align-items: flex-start; padding: 6px 9px;
            border: 1px solid var(--line); border-radius: 7px; background: #fff; }
        .epf-mrow.epf-facing .epf-fcell { --w: 1 1 0; }
        .epf-fcell.epf-fwcell { --w: 1 1 100%; }
        .epf-fcell[data-key="paymentComment"] { --w: 0 0 calc(50% - 4px); max-width: calc(50% - 4px); }
        .epf-fcell[data-key="nationalId"], .epf-fcell[data-key="llManner"] { --w: 0.8 1 0; }
        .epf-fcell[data-key="phoneNumber"] { --w: 1.4 1 0; }
        .epf-fcell[data-key="cancellationDate"] { --w: 0 0 285px; max-width: 36%; }
        .epf-fcell[data-key="cancellationComment"] { --w: 1 1 0; }
        .epf-fcell[data-st="will-change"] { background: #f2faf3; border-color: #b9dfc0; }
        .epf-fcell[data-st="same"] { background: #f4f6f7; border-color: #d8dedf; }
        .epf-fcell[data-st="not-editable"] { background: #fffbeb; border-color: #f0d78c; }
        .epf-fcell[data-st="invalid-select"] { background: #fff4f4; border-color: #efb7bd; }
        #epf-fields.epf-only-changes .epf-fcell:not([data-st="will-change"]):not([data-unmapped="1"]) { display: none; }
        .epf-flabel { flex: 0 0 auto; max-width: 45%; min-width: 82px; padding-top: 5px; direction: ltr; text-align: left; color: #34464d;
            font-size: 12.5px; font-weight: 700; overflow-wrap: anywhere; }
        .epf-flabel bdi::after { content: ":"; margin-inline-start: 1px; }
        .epf-fwrap { flex: 1 1 auto; width: 100%; min-width: 0; direction: rtl; }
        .epf-field-control {
            display: block; width: 100%; min-height: 32px; padding: 5px 7px; border: 1px solid #cbd3d6;
            border-radius: 6px; color: var(--ink); background: #fff; font-size: 12.5px; line-height: 1.5; text-align: start;
        }
        .epf-field-control:disabled { color: #7e8b90; background: #f1f3f4; }
        select.epf-field-control { direction: ltr; text-align: left; }
        textarea.epf-field-control { min-height: 44px; resize: vertical; white-space: pre-wrap; }
        .epf-meta { margin-top: 3px; font-size: 11.5px; line-height: 1.45; color: var(--mut); }
        .epf-field-status { display: block; min-height: 15px; }
        .epf-field-cur { display: block; max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #4e5e64; font-size: 11.5px; margin-top: 1px; }
        [data-st="will-change"] .epf-field-status { color: var(--ok); font-weight: 700; }
        [data-st="not-editable"] .epf-field-status { color: var(--warn); }
        [data-st="invalid-select"] .epf-field-status { color: var(--bad); font-weight: 700; }
        .epf-ct { margin-top: 5px; }
        .epf-ct-btn {
            display: inline-flex; align-items: center; gap: 5px; padding: 5px 10px; border: 1px solid var(--acc);
            border-radius: 6px; color: var(--acc); background: var(--pri); font-size: 12px; font-weight: 700; cursor: pointer;
        }
        .epf-ct-btn:hover:not(:disabled) { background: var(--pri-d); }
        .epf-ct-btn:disabled { color: #8b979c; background: #d9dfe1; border-color: #c3cbce; cursor: not-allowed; }
        .epf-ct-msg { margin-top: 4px; font-size: 11.5px; line-height: 1.5; }
        .epf-ct-info { color: #7a5b00; }
        .epf-ct-bad { color: var(--bad); }
        .epf-ct-ok { color: var(--ok); font-weight: 700; }

        /* Sticky footer keeps Apply/Cancel always reachable */
        .epf-footer {
            position: sticky; bottom: 0; z-index: 2;
            display: flex; flex-direction: column; gap: 7px;
            padding: 8px 13px 10px;
            border-top: 1px solid #e8edef; background: #fbfcfc;
            box-shadow: 0 -2px 6px rgba(0,0,0,.04);
        }
        .epf-settings-footer { direction: rtl; display: flex; flex-direction: row; justify-content: center; align-items: center; gap: 14px; flex-wrap: wrap; }
        .epf-settings-footer .epf-btn { min-width: 145px; min-height: 38px; padding: 7px 16px; }
        .epf-settings-footer .epf-btn-reset { min-width: 125px; }
        .epf-chips { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; margin-inline-end: 16px; }
        .epf-chip { display: inline-flex; align-items: center; gap: 5px; padding: 3px 9px; border-radius: 20px; background: #edf0f1; color: #4b585f; font-size: 11.5px; white-space: nowrap; }
        .epf-chip::before { content: ''; width: 7px; height: 7px; border-radius: 50%; background: #9aa5aa; }
        .epf-chip[data-k="will-change"]::before { background: var(--ok); }
        .epf-chip[data-k="invalid-select"]::before { background: var(--bad); }
        .epf-chip[data-k="not-editable"]::before { background: #d2a500; }
        .epf-chip-count { margin-inline-start: 2px; font-size: 14px; line-height: 1; font-weight: 800; color: inherit; }
        .epf-only { display: inline-flex; align-items: center; gap: 7px; font-size: 12.5px; font-weight: 700; color: var(--ink); white-space: nowrap; margin-inline-start: auto; cursor: pointer; }
        .epf-only input { width: 17px; height: 17px; accent-color: var(--pri); cursor: pointer; }
        .epf-footer-actions { display: flex; flex-wrap: wrap; align-items: center; gap: 12px; width: 100%; }

        #epf-toast {
            position: fixed; left: 24px; bottom: 20px; z-index: 10002; max-width: min(92vw, 560px); padding: 12px 16px;
            border-radius: 10px; color: #fff; background: var(--pri); box-shadow: 0 6px 22px rgba(0,0,0,.22);
            font-size: 12.5px; line-height: 1.8; direction: rtl; cursor: pointer; opacity: 0; transform: translateY(8px);
            pointer-events: none; transition: .2s ease;
        }
        #epf-toast.epf-toast-show { opacity: 1; transform: translateY(0); pointer-events: auto; }
        .epf-filled { background-color: #e8f5e9 !important; border-color: #28a745 !important; }

        .epf-settings { flex: 1 1 auto; min-height: 0; overflow: visible; }
        .epf-settings-topbar { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 6px; }
        .epf-settings-intro { margin: 0 0 9px; font-size: 13px; line-height: 1.85; color: #36464d; text-align: right; }
        .epf-settings-layout { display: grid; grid-template-columns: minmax(0, 1fr) 138px; gap: 8px; direction: ltr; align-items: stretch; }
        .epf-settings-actions { display: flex; flex-direction: column; gap: 6px; }
        .epf-settings-actions .epf-btn { min-height: 36px; padding: 6px 8px; font-size: 12px; }
        .epf-settings-area { display: block; width: 100%; min-height: 62px; max-height: 135px; resize: vertical; padding: 7px 10px;
            border: 1px solid #cbd3d6; border-radius: 7px; background: #fbfcfc; color: var(--ink); font-size: 12.5px;
            line-height: 1.55; direction: ltr; text-align: left; white-space: pre; }
        .epf-hlegend { display: flex; flex-wrap: wrap; align-items: baseline; gap: 6px 18px; margin: 10px 0 0; font-size: 12px; line-height: 1.65; color: var(--mut); }
        .epf-hlegend span { font-weight: 700; }
        .epf-hlegend .lg-miss { color: var(--bad); }
        .epf-hlegend .lg-dup { color: #c2570c; }
        .epf-hlegend .lg-abs { color: #9a7b0a; }
        .epf-hlegend .lg-chg { color: var(--ok); }
        .epf-hlegend .lg-def { color: #8a959b; }
        .epf-hlegend .lg-dis { color: #8b979c; text-decoration: line-through; }
        .epf-hgrid-wrap { margin-top: 9px; padding: 10px; border: 1px solid var(--line); border-radius: 8px; background: #fbfcfc; }
        .epf-hgrid-site-row { display: flex; justify-content: center; margin-bottom: 9px; }
        .epf-hgrid-site-row .epf-hcell { width: min(270px, 100%); border-color: #b8dce8; background: #f2f9fc; }
        .epf-hgrid { direction: ltr; display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 7px; }
        .epf-hcell { display: flex; align-items: center; justify-content: space-between; gap: 6px; min-width: 0; padding: 7px 9px;
            border: 1px solid #e2e8ea; border-radius: 6px; background: #fff; font-size: 12.5px; }
        .epf-hchk { width: 15px; height: 15px; accent-color: var(--pri); cursor: pointer; flex: 0 0 auto; margin: 0; }
        .epf-hcell-left { display: flex; align-items: center; gap: 6px; min-width: 0; flex: 1 1 auto; }
        .epf-hname { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ink); font-weight: 600; }
        .epf-hname.epf-h-disabled { opacity: 0.55; text-decoration: line-through; color: #7f8c8d; }
        .epf-hcol { flex: 0 0 auto; min-width: 18px; text-align: center; font-weight: 800; font-size: 13px; color: #98a3a8; }
        .epf-hname.epf-h-missing, .epf-hcol.epf-h-missing { color: var(--bad); }
        .epf-hname.epf-h-dup, .epf-hcol.epf-h-dup { color: #c2570c; }
        .epf-hname.epf-h-absent, .epf-hcol.epf-h-absent { color: #9a7b0a; }
        .epf-hcol.epf-h-changed { color: var(--ok); }
        .epf-box { display: none; margin-top: 9px; padding: 8px 10px; border-radius: 7px; font-size: 12px; line-height: 1.75; white-space: pre-wrap; }
        .epf-box.epf-show { display: block; }
        .epf-box-err { color: #842029; background: #f8d7da; }
        .epf-box-warn { color: #7a5b00; background: #fff8e6; border: 1px solid #f0d78c; }
        .epf-box-info { color: #174b58; background: #e8f1f3; }
        .epf-box-ok { color: #155724; background: #d4edda; }

        #epf-confirm-overlay { position: fixed; inset: 0; z-index: 10005; display: none; align-items: center; justify-content: center;
            padding: 16px; background: rgba(30,25,21,.45); }
        #epf-confirm-overlay.epf-open { display: flex; }
        .epf-confirm-box { width: min(92vw, 430px); padding: 18px 20px; background: #fff; color: var(--ink); border-radius: 10px;
            box-shadow: 0 14px 44px rgba(0,0,0,.3); }
        .epf-confirm-msg { margin: 0 0 15px; font-size: 13px; line-height: 1.85; }
        .epf-confirm-actions { display: flex; justify-content: center; gap: 10px; }
        .epf-confirm-actions .epf-btn { min-width: 135px; }

        @media (max-width: 768px) {
            #epf-header-btn { padding: 6px 10px; font-size: 12.5px; gap: 4px; }
            .epf-content { padding: 8px; }
            .epf-mrow { flex-wrap: wrap; }
            .epf-fcell { --w: 1 1 100% !important; flex: 1 1 100% !important; max-width: 100% !important; }
            .epf-flabel { max-width: 100%; flex: 0 0 auto; }
            .epf-settings-footer { flex-wrap: wrap; }
            .epf-footer { padding: 7px 8px; }
            .epf-hgrid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            .epf-settings-layout { grid-template-columns: minmax(0, 1fr) 115px; }
            .epf-identity { max-width: 180px; font-size: 11.5px; }
        }
        @media (max-width: 420px) {
            .epf-title { font-size: 14px; }
            .epf-subtitle { font-size: 11px; }
            .epf-paste-layout { grid-template-columns: minmax(0, 1fr) 112px; }
            #epf-paste-area { min-height: 46px; }
            .epf-chip { font-size: 10.5px; padding: 2px 6px; }
            .epf-chip-count { font-size: 13px; }
        }
        @media (prefers-reduced-motion: reduce) {
            #epf-overlay { transition: none; }
            #epf-header-btn, .epf-fcell { transition: none; }
        }
    `;

    // ============================================================
    // 05. INLINE PANEL UI
    // ============================================================
    const root = makeElement('div');
    root.id = 'epf-root';

    const overlay = makeElement('div');
    overlay.id = 'epf-overlay';
    overlay.setAttribute('aria-hidden', 'true');

    const modal = makeElement('section');
    modal.id = 'epf-modal';
    modal.setAttribute('role', 'region');
    modal.setAttribute('aria-labelledby', 'epf-title');

    const header = makeElement('div', 'epf-header');
    const headerTop = makeElement('div', 'epf-header-top');
    const title = makeElement('h2', 'epf-title', 'انتقال اطلاعات Excel به فرم BTS Property');
    title.id = 'epf-title';
    const gearBtn = makeElement('button', 'epf-gear');
    gearBtn.type = 'button';
    gearBtn.setAttribute('aria-label', 'تنظیمات نگاشت هدر');
    const gearState = makeElement('span', 'epf-gear-state', 'پیش‌فرض');
    gearBtn.append(makeElement('span', 'epf-gear-main', '⚙ تنظیمات نگاشت هدر'), gearState);

    const headerIdentity = makeElement('span', 'epf-identity', '— —');
    headerIdentity.id = 'epf-identity';
    headerIdentity.setAttribute('role', 'status');

    const closeMainBtn = makeElement('button', 'epf-xbtn', '✕');
    closeMainBtn.id = 'epf-x-main';
    closeMainBtn.type = 'button';
    closeMainBtn.setAttribute('aria-label', 'بستن پنل');
    const closeSettingsBtn = makeElement('button', 'epf-xbtn', '✕');
    closeSettingsBtn.id = 'epf-x-settings';
    closeSettingsBtn.type = 'button';
    closeSettingsBtn.setAttribute('aria-label', 'بستن تنظیمات');
    closeSettingsBtn.hidden = true;

    const headerMain = makeElement('div');
    const subtitle = makeElement('p', 'epf-subtitle',
        'ردیف داده را از اکسل در کادر زیر پیست کنید و «بررسی» را بزنید. فیلدهای بدون هدر یا بدون داده بررسی نمی‌شوند و خانهٔ خالی یا حاوی خطای فرمول اکسل اعمال نمی‌شود. اگر فیلدی مؤثر در محاسبه تغییر کند، Calculate خودکار اجرا می‌شود.');
    const mapNote = makeElement('div');
    mapNote.id = 'epf-map-note';
    mapNote.setAttribute('role', 'status');

    const pasteLayout = makeElement('div', 'epf-paste-layout');
    const pasteArea = makeElement('textarea');
    pasteArea.id = 'epf-paste-area';
    pasteArea.placeholder = 'ردیف داده را اینجا پیست کنید یا از دکمه «پیست» استفاده کنید...';
    pasteArea.spellcheck = false;
    pasteArea.setAttribute('aria-label', 'ردیف داده اکسل');
    const pasteActions = makeElement('div', 'epf-paste-actions');
    const pasteButton = makeElement('button', 'epf-btn epf-btn-secondary', 'پیست');
    pasteButton.type = 'button';
    const reviewButton = makeElement('button', 'epf-btn epf-btn-primary', 'بررسی');
    reviewButton.type = 'button';
    pasteActions.append(pasteButton, reviewButton);
    pasteLayout.append(pasteArea, pasteActions);

    const pasteWarning = makeElement('div');
    pasteWarning.id = 'epf-paste-warning';
    pasteWarning.setAttribute('role', 'status');
    pasteWarning.setAttribute('aria-live', 'polite');
    headerMain.append(subtitle, mapNote, pasteLayout, pasteWarning);
    const titleWrap = makeElement('div', 'epf-title-wrap');
    titleWrap.append(closeMainBtn, closeSettingsBtn, title);
    headerTop.append(titleWrap, headerIdentity, gearBtn);
    header.append(headerTop, headerMain);

    const content = makeElement('div', 'epf-content');
    const mainMessage = makeElement('div');
    mainMessage.id = 'epf-main-message';
    mainMessage.setAttribute('role', 'status');
    mainMessage.setAttribute('aria-live', 'polite');
    const fieldsContainer = makeElement('div');
    fieldsContainer.id = 'epf-fields';
    content.append(mainMessage, fieldsContainer);

    const footer = makeElement('div', 'epf-footer');
    const chipsBox = makeElement('div', 'epf-chips');
    const CHIP_DEFS = [
        { key: 'will-change', label: 'تغییر', hint: 'فیلدهایی که مقدارشان روی فرم تغییر می‌کند' },
        { key: 'same', label: 'یکسان / خالی', hint: 'مقدار ردیف با فرم برابر است یا خانهٔ ردیف خالی / خطای اکسل است' },
        { key: 'invalid-select', label: 'نامعتبر', hint: 'مقدار ردیف در گزینه‌های لیست سایت نیست و اعمال نمی‌شود' },
        { key: 'not-editable', label: 'غیرقابل بررسی', hint: 'بدون هدر، بدون داده در ردیف یا غیرقابل ویرایش در فرم؛ بررسی نمی‌شود' }
    ];
    const chips = CHIP_DEFS.map(def => {
        const el = makeElement('span', 'epf-chip');
        el.dataset.k = def.key;
        el.title = def.hint;
        el.append(makeElement('span', '', def.label), makeElement('strong', 'epf-chip-count', '0'));
        chipsBox.appendChild(el);
        return { key: def.key, label: def.label, el, count: el.querySelector('.epf-chip-count') };
    });
    const onlyLabel = makeElement('label', 'epf-only');
    const onlyBox = makeElement('input');
    onlyBox.type = 'checkbox';
    onlyLabel.append(onlyBox, document.createTextNode('فقط تغییرات'));
    const footerActions = makeElement('div', 'epf-footer-actions');
    const applyButton = makeElement('button', 'epf-btn epf-btn-apply', 'اعمال در فرم وب');
    applyButton.type = 'button';
    applyButton.disabled = true;
    const gotoSaveButton = makeElement('button', 'epf-btn epf-btn-goto-save', '⬇️ رفتن به دکمه ذخیره');
    gotoSaveButton.type = 'button';
    gotoSaveButton.id = 'epf-goto-save';
    const cancelButton = makeElement('button', 'epf-btn epf-btn-cancel', 'انصراف');
    cancelButton.type = 'button';
    footerActions.append(chipsBox, onlyLabel, applyButton, gotoSaveButton, cancelButton);
    footer.append(footerActions);

    const settingsView = makeElement('div', 'epf-content epf-settings');
    settingsView.hidden = true;
    const settingsFooter = makeElement('div', 'epf-footer epf-settings-footer');
    settingsFooter.hidden = true;
    modal.append(header, content, settingsView, footer, settingsFooter);
    overlay.appendChild(modal);
    root.appendChild(overlay);

    const toast = makeElement('div');
    toast.id = 'epf-toast';
    toast.setAttribute('role', 'status');
    toast.setAttribute('aria-live', 'polite');
    root.appendChild(toast);

    const confirmOverlay = makeElement('div');
    confirmOverlay.id = 'epf-confirm-overlay';
    confirmOverlay.setAttribute('aria-hidden', 'true');
    const confirmBox = makeElement('div', 'epf-confirm-box');
    confirmBox.setAttribute('role', 'alertdialog');
    confirmBox.setAttribute('aria-modal', 'true');
    const confirmMsg = makeElement('p', 'epf-confirm-msg');
    const confirmOkBtn = makeElement('button', 'epf-btn epf-btn-primary', 'برگشت بدون ذخیره');
    confirmOkBtn.type = 'button';
    const confirmStayBtn = makeElement('button', 'epf-btn epf-btn-secondary', 'ماندن در تنظیمات');
    confirmStayBtn.type = 'button';
    const confirmActions = makeElement('div', 'epf-confirm-actions');
    confirmActions.append(confirmOkBtn, confirmStayBtn);
    confirmBox.append(confirmMsg, confirmActions);
    confirmOverlay.appendChild(confirmBox);
    root.appendChild(confirmOverlay);

    // ============================================================
    // 06. FIELD REVIEW CONTROLS
    // ============================================================

    function isMapped(key) {
        return Boolean(isFieldEnabled(key) && activeHeaderMapping && Number.isInteger(activeHeaderMapping.columns[key]));
    }

    function isCancellationDateEnabled() {
        if (String(getVal('cancellationLetter')).trim().toUpperCase() === 'YES') return true;
        const info = targetInfo(FIELD_BY_KEY.cancellationLetter);
        return Boolean(info.exists && info.typeMatches &&
            String(info.element.value || '').trim().toUpperCase() === 'YES');
    }

    function buildRows() {
        if (built) return;
        built = true;
        const fragment = document.createDocumentFragment();

        MODAL_ROWS.forEach(section => {
            const group = GROUPS.find(item => item.id === section.id);
            const box = makeElement('section', 'epf-section');
            box.appendChild(makeElement('div', 'epf-section-title', group ? group.title : ''));
            groupEls.push({ el: box, id: section.id });

            section.rows.forEach(rowDef => {
                const rowEl = makeElement('div', 'epf-mrow' + (rowDef.facing ? ' epf-facing' : ''));
                rowDef.fields.forEach(key => {
                    const field = FIELD_BY_KEY[key];
                    const cell = makeElement('div', 'epf-fcell' + (field.type === 'textarea' ? ' epf-fwcell' : ''));
                    cell.dataset.key = key;
                    cell.dataset.st = 'idle';

                    const label = makeElement('label', 'epf-flabel');
                    const name = makeElement('bdi', '', field.label);
                    name.dir = 'ltr';
                    label.appendChild(name);

                    const wrap = makeElement('div', 'epf-fwrap');
                    const controlType = field.type === 'select' ? 'select'
                        : field.type === 'textarea' ? 'textarea' : 'input';
                    const control = makeElement(controlType, 'epf-field-control');
                    control.id = `epf-control-${key}`;
                    control.dataset.key = key;
                    if (controlType === 'input') {
                        control.type = 'text';
                        control.dir = 'auto';
                        control.autocomplete = 'off';
                        control.spellcheck = false;
                    } else if (controlType === 'textarea') {
                        control.rows = 2;
                        control.dir = 'auto';
                        control.spellcheck = false;
                    }
                    label.htmlFor = control.id;

                    const meta = makeElement('div', 'epf-meta');
                    const status = makeElement('span', 'epf-field-status', 'پس از بررسی ردیف نمایش داده می‌شود');
                    const current = makeElement('span', 'epf-field-cur');
                    current.hidden = true;
                    const currentValue = makeElement('bdi');
                    currentValue.dir = 'auto';
                    current.append(document.createTextNode('فعلی در وب: '), currentValue);
                    meta.append(status, current);
                    wrap.append(control, meta);
                    if (key === FIELD_KEYS.CONTRACT_TYPE) wrap.appendChild(buildContractBox());
                    cell.append(label, wrap);
                    rowEl.appendChild(cell);

                    controls[key] = control;
                    rows[key] = { row: cell, status, cur: current, curVal: currentValue };
                    control.addEventListener(field.type === 'select' ? 'change' : 'input', () => {
                        if (state.excelErrors && state.excelErrors[key]) {
                            delete state.excelErrors[key];
                        }
                        state.webValues[key] = control.value;
                        refreshAll();
                    });
                });
                box.appendChild(rowEl);
                modalRowEls.push({ el: rowEl, keys: rowDef.fields });
            });
            fragment.appendChild(box);
        });

        fieldsContainer.appendChild(fragment);
        WEB_FIELDS.forEach(field => syncControl(field.key));
    }

    function syncControl(key) {
        const control = controls[key];
        if (!control) return;
        const field = FIELD_BY_KEY[key];
        const value = getVal(key);
        const isOut = Boolean(state.outOfRange && state.outOfRange[key]);
        if (!isFieldEnabled(key) || !state.parsed || !isMapped(key) || isOut) {
            control.disabled = true;
            if (field.type === 'select') {
                control.textContent = '';
                const placeholderText = !isFieldEnabled(key) ? '(غیرفعال)'
                    : !state.parsed ? '(پس از بررسی)'
                    : !isMapped(key) ? '(هدر ندارد)' : '(داده ندارد)';
                const option = makeElement('option', '', placeholderText);
                option.value = '';
                control.appendChild(option);
            } else control.value = '';
            return;
        }

        if (field.type !== 'select') {
            control.value = field.type === 'textarea' ? value : String(value).replace(/\r?\n/g, ' ');
            control.disabled = key === 'cancellationDate' && !isCancellationDateEnabled();
            return;
        }

        const info = targetInfo(field);
        const match = info.exists && info.typeMatches ? findSelectableOption(info.element, value) : null;
        control.textContent = '';
        if (!match) {
            const extra = makeElement('option', '', value === '' ? '(خالی)' : '⚠ ' + value);
            extra.value = value;
            extra.selected = true;
            control.appendChild(extra);
        }
        if (info.exists && info.typeMatches) {
            Array.from(info.element.options).forEach(option => {
                const item = makeElement('option', '', String(option.textContent || '').trim());
                item.value = option.value;
                item.disabled = option.disabled;
                if (match === option) item.selected = true;
                control.appendChild(item);
            });
        }
        control.disabled = false;
    }

    function syncAllSelects() {
        WEB_FIELDS.forEach(field => {
            if (field.type === 'select') syncControl(field.key);
        });
        syncControl('cancellationDate');
    }

    // ============================================================
    // 07. FIELD STATUS & COMPARISON
    // ============================================================
    function showMainMessage(message, kind) {
        mainMessage.textContent = message || '';
        const className = kind === 'error' ? 'epf-msg-error'
            : kind === 'success' ? 'epf-msg-success'
            : kind === 'warning' ? 'epf-msg-warning' : 'epf-msg-info';
        mainMessage.className = 'epf-show ' + className;
    }

    function clearMainMessage() {
        mainMessage.textContent = '';
        mainMessage.className = '';
    }

    function setRowStatus(key, statusType, message, currentText, reason) {
        const item = rows[key];
        if (!item) return;
        item.row.dataset.st = statusType;
        if (reason) {
            item.row.dataset.reason = reason;
            if (reason === 'no-header' || reason === 'no-data') item.row.dataset.unmapped = '1';
            else delete item.row.dataset.unmapped;
        } else {
            delete item.row.dataset.reason;
            delete item.row.dataset.unmapped;
        }
        item.status.textContent = message || '';
        if (currentText === undefined || currentText === null) {
            item.cur.hidden = true;
        } else {
            item.curVal.textContent = currentText === '' ? '(خالی)' : currentText;
            item.cur.title = currentText;
            item.cur.hidden = false;
        }
    }

    function updateChips(counts) {
        chips.forEach(chip => {
            chip.count.textContent = state.parsed && counts ? String(counts[chip.key] || 0) : '0';
        });
    }

    function updateGroups() {
        const onlyChanges = onlyBox.checked;
        fieldsContainer.classList.toggle('epf-only-changes', onlyChanges);
        const isVisibleInOnly = key => isFieldEnabled(key) && rows[key] && (rows[key].row.dataset.st === 'will-change' || rows[key].row.dataset.unmapped === '1');
        modalRowEls.forEach(item => {
            const hasVisibleKeys = item.keys.some(k => isFieldEnabled(k) && (!onlyChanges || isVisibleInOnly(k)));
            item.el.hidden = !hasVisibleKeys;
        });
        groupEls.forEach(group => {
            const hasVisible = WEB_FIELDS.some(field => field.group === group.id && isFieldEnabled(field.key) && (!onlyChanges || isVisibleInOnly(field.key)));
            group.el.hidden = !hasVisible;
        });
    }

    function evaluateField(key) {
        const field = FIELD_BY_KEY[key];
        if (!isFieldEnabled(key)) {
            if (rows[key]) rows[key].row.hidden = true;
            return 'disabled';
        }
        if (rows[key]) rows[key].row.hidden = false;

        if (!isMapped(key)) {
            setRowStatus(key, 'not-editable', 'هدر ندارد؛ بررسی نمی‌شود', null, 'no-header');
            return 'not-editable';
        }
        if (state.outOfRange && state.outOfRange[key]) {
            setRowStatus(key, 'not-editable', 'داده ندارد؛ بررسی نمی‌شود', null, 'no-data');
            return 'not-editable';
        }

        const current = getCurrentFormValue(field);
        const pasted = getVal(key);
        const rawExcelError = state.excelErrors ? state.excelErrors[key] : null;

        if (String(pasted) === '') {
            const emptyMsg = rawExcelError
                ? `خطای اکسل (${rawExcelError}) نادیده گرفته شد؛ اعمال نمی‌شود`
                : 'داده خالی است؛ اعمال نمی‌شود';
            setRowStatus(key, 'same', emptyMsg, current);
            return 'same';
        }

        const info = targetInfo(field);

        if (key === 'cancellationDate') {
            if (!isCancellationDateEnabled()) {
                setRowStatus(key, 'not-editable', 'فقط وقتی Cancellation Letter برابر Yes باشد قابل اعمال است', current);
                return 'not-editable';
            }
        } else {
            if (!info.exists || !info.typeMatches) {
                setRowStatus(key, 'not-editable', 'فیلد در فرم پیدا نشد یا نوع آن سازگار نیست؛ بررسی نمی‌شود');
                return 'not-editable';
            }
            if (!info.editable) {
                const typeRow = rows[FIELD_KEYS.ESCALATION_TYPE];
                if (key === 'escalation' && typeRow && typeRow.row.dataset.st === 'will-change') {
                    setRowStatus(key, 'not-editable', 'پس از اعمال Escalation Type باز می‌شود و همان‌جا اعمال خواهد شد', current, 'deferred');
                    return 'not-editable';
                }
                setRowStatus(key, 'not-editable', 'این فیلد در فرم قابل ویرایش نیست؛ اعمال نمی‌شود', current);
                return 'not-editable';
            }
        }

        let option = null;
        if (field.type === 'select') {
            option = findSelectableOption(info.element, pasted);
            if (!option) {
                setRowStatus(key, 'invalid-select', 'مقدار در گزینه‌های سایت وجود ندارد؛ اعمال نمی‌شود', current);
                return 'invalid-select';
            }
        }

        let same;
        if (field.type === 'select') same = String(info.element.value) === String(option.value);
        else if (key === FIELD_KEYS.AMOUNT) same = stripCommas(pasted) === stripCommas(current || '');
        else same = looseText(pasted) === looseText(current || '');

        if (same) {
            setRowStatus(key, 'same', 'مقدار ردیف با فرم یکسان است؛ تغییری لازم نیست');
            return 'same';
        }
        setRowStatus(key, 'will-change', 'مقدار فرم تغییر خواهد کرد', current);
        return 'will-change';
    }

    function validateIdentity() {
        const siteInfo = targetInfo(FIELD_BY_KEY[FIELD_KEYS.SITE_ID]);
        const contractInfo = targetInfo(FIELD_BY_KEY[FIELD_KEYS.CONTRACT_TYPE]);
        const result = { ok: false, siteOk: false, contractOk: false, message: '' };

        if (!state.parsed) {
            result.message = 'ابتدا یک ردیف معتبر وارد کرده و بررسی کنید.';
            return result;
        }
        if (!state.searchConfirmed) {
            result.message = 'جستجوی Site ID هنوز تأیید نشده است؛ اطلاعات تا تکمیل جستجو قابل اعمال نیست.';
            return result;
        }

        if (!siteInfo.exists || !siteInfo.typeMatches || !contractInfo.exists || !contractInfo.typeMatches) {
            result.message = 'فیلد Site ID یا Contract Type در صفحه پیدا نشد.';
            return result;
        }

        const pastedSite = normalize(getVal(FIELD_KEYS.SITE_ID));
        const pageSite = normalize(siteInfo.element.value);
        const pastedContract = normalize(getVal(FIELD_KEYS.CONTRACT_TYPE));
        const pageContract = normalize(selectedOptionText(contractInfo.element));

        result.siteOk = Boolean(pastedSite && pageSite && pastedSite === pageSite);
        result.contractOk = Boolean(pastedContract && pageContract && pastedContract === pageContract);
        result.ok = result.siteOk && result.contractOk;

        if (result.ok) {
            result.message = 'Site ID و Contract Type با فرم فعلی مطابقت دارند.';
        } else {
            const details = [];
            if (!result.siteOk) details.push('Site ID منطبق نیست یا خالی است');
            if (!result.contractOk) details.push('Contract Type منطبق نیست یا خالی است');
            result.message = 'Site ID یا Contract Type مطابقت ندارد: ' + details.join('؛ ');
        }
        return result;
    }

    function contractDiffersFromForm() {
        const info = targetInfo(FIELD_BY_KEY[FIELD_KEYS.CONTRACT_TYPE]);
        if (!info.exists || !info.typeMatches) return false;
        const pasted = normalize(getVal(FIELD_KEYS.CONTRACT_TYPE));
        return Boolean(pasted) && pasted !== normalize(selectedOptionText(info.element));
    }

    function refreshAll() {
        if (!built) return;
        syncControl('cancellationDate');
        if (!state.parsed) {
            updateChips(null);
            updateGroups();
            applyButton.disabled = true;
            return;
        }

        const counts = { 'will-change': 0, 'same': 0, 'not-editable': 0, 'invalid-select': 0 };
        WEB_FIELDS.forEach(field => {
            if (!isFieldEnabled(field.key)) {
                if (rows[field.key]) rows[field.key].row.hidden = true;
                return;
            }
            const st = evaluateField(field.key);
            if (st !== 'disabled' && counts[st] !== undefined) counts[st] += 1;
        });
        updateChips(counts);
        updateGroups();

        const identity = validateIdentity();
        updateApplyState();

        if (state.searching || state.calculating) return;
        const invalidFields = WEB_FIELDS.filter(field => isFieldEnabled(field.key) && rows[field.key].row.dataset.st === 'invalid-select');
        let message = identity.message;
        let kind = identity.ok ? 'success' : 'error';
        if (invalidFields.length) {
            const labels = invalidFields.map(field => field.label).join('، ');
            if (contractDiffersFromForm()) {
                message += '\n' + faNum(invalidFields.length) + ' مقدار Dropdown فعلاً با گزینه‌های سایت تطبیق ندارد: ' + labels +
                    '. چون نوع قرارداد داده با فرم فرق دارد، ابتدا «اعمال Contract جدید» را بزنید؛ ' +
                    'اگر بعد از اعمال هم نامعتبر ماند، مشکل واقعی ستون‌هاست.';
                kind = 'info';
            } else if (invalidFields.length > 1) {
                message += '\nچند مقدار Dropdown با گزینه‌های سایت تطبیق ندارد: ' + labels +
                    '. احتمال جابه‌جایی ستون‌ها وجود دارد؛ هدر ثبت‌شده و ردیف اکسل را بررسی کنید.';
                kind = 'error';
            } else {
                message += '\nمقدار Dropdown این فیلد در گزینه‌های سایت وجود ندارد: ' + labels +
                    '. ستون متناظر را بررسی کنید.';
                kind = 'warning';
            }
        }
        showMainMessage(message, kind);
    }

    // ============================================================
    // 08. SITE ID SEARCH
    // ============================================================
    function siteIdMissing() {
        return Boolean(document.getElementById(SITE_INSERT_INPUT_ID) || document.getElementById('Content_Site'));
    }

    async function performAutoSearch(siteId) {
        const siteInput = document.getElementById(SITE_ID_INPUT_ID);
        const searchButton = document.getElementById(SEARCH_BTN_ID);
        if (!siteInput || !searchButton) {
            return { ok: false, reason: 'فیلد Site ID یا دکمه جستجو در صفحه پیدا نشد.', kind: 'error' };
        }

        siteInput.value = siteId;
        siteInput.dispatchEvent(new Event('input', { bubbles: true }));
        siteInput.dispatchEvent(new Event('change', { bubbles: true }));
        showMainMessage('در حال جستجوی Site ID...', 'info');

        const slowTimer = setTimeout(() => {
            showMainMessage('جستجو همچنان در حال اجرا است؛ چند لحظه صبر کنید.', 'info');
        }, 4000);

        state.cancelRequested = false;
        const result = await runAndWait(() => searchButton.click(), 2000, 15000, null, { abortOnCancel: true });
        clearTimeout(slowTimer);

        const cancelledResult = { ok: false, kind: 'warning',
            reason: 'انتظار جستجو را لغو کردید؛ وضعیت فرم را بررسی کنید و برای ادامه دوباره «بررسی» را بزنید.' };
        if (result.cancelled || state.cancelRequested) return cancelledResult;

        if (result.error) {
            return { ok: false, reason: 'هنگام جستجوی Site ID خطا رخ داد. وضعیت صفحه را بررسی کنید.', kind: 'error' };
        }
        if (result.timeout) {
            return { ok: false, reason: 'عملیات جستجو در زمان مورد انتظار تکمیل نشد؛ احتمال اختلال ارتباط یا پایان نشست کاربری وجود دارد.', kind: 'warning' };
        }
        if (siteIdMissing()) {
            return { ok: false, reason: 'این Site ID در سامانه ثبت نشده است؛ فرم ثبت سایت جدید نمایش داده شده است.', kind: 'error' };
        }

        return { ok: true, reason: '', kind: 'info' };
    }

    // ============================================================
    // 09. CLIPBOARD & ROW REVIEW
    // ============================================================
    function isPasteStale() {
        return state.parsed && pasteArea.value !== state.reviewedText;
    }

    function updatePasteRevisionWarning() {
        const changed = isPasteStale();
        pasteWarning.textContent = changed
            ? 'متن کادر پس از بررسی تغییر کرده است؛ تا بررسی دوباره، دکمهٔ «اعمال در فرم وب» غیرفعال است.'
            : '';
        pasteWarning.classList.toggle('epf-show', changed);
        if (built) updateApplyState();
    }

    function updateApplyState() {
        const busy = state.searching || state.calculating;
        const hasChanges = WEB_FIELDS.some(field => isFieldEnabled(field.key) && rows[field.key] && rows[field.key].row.dataset.st === 'will-change');
        applyButton.disabled = busy || !state.parsed || isPasteStale() || !hasChanges || !validateIdentity().ok;
    }

    function setBusy(busy) {
        pasteButton.disabled = busy;
        reviewButton.disabled = busy || !activeHeaderMapping;
        pasteArea.disabled = busy;
        gearBtn.disabled = busy;
        if (ctBtn) ctBtn.disabled = busy;
        cancelButton.textContent = !busy ? 'انصراف' : (state.calculating ? 'توقف پیگیری' : 'لغو انتظار');
        cancelButton.disabled = false;
        updateApplyState();
    }

    function resetParsedData() {
        state.webValues = {};
        state.excelErrors = {};
        state.outOfRange = {};
        state.parsed = false;
        state.reviewedText = '';
        state.searchConfirmed = false;
        state.headerContract = '';
        state.contractChangeCompleted = false;
        if (contractFixTimer) { clearTimeout(contractFixTimer); contractFixTimer = null; }
        updateHeaderIdentity();
        WEB_FIELDS.forEach(field => {
            syncControl(field.key);
            setRowStatus(field.key, 'idle', 'پس از بررسی ردیف نمایش داده می‌شود');
        });
        applyButton.disabled = true;
        resetContractFix();
        updateChips(null);
        updateGroups();
        updatePasteRevisionWarning();
    }

    function processPastedText(raw) {
        if (!activeHeaderMapping) {
            return { ok: false, message: 'ابتدا هدر اکسل را در «تنظیمات نگاشت هدر» ثبت کنید.' };
        }
        const allRows = parseTSV(raw);
        if (!allRows.length) return { ok: false, message: 'متنی برای بررسی وجود ندارد.' };

        const cells = allRows[0];
        const extraction = extractValues(cells, activeHeaderMapping);
        if (!extraction.ok) return extraction;

        const rowWarning = allRows.length > 1
            ? `توجه: ${allRows.length} ردیف تشخیص داده شد؛ فقط ردیف اول بررسی شد. ` : '';
        const lengths = allRows.map(row => row.length);
        const colWarning = allRows.length > 1 && lengths.some(n => n !== lengths[0])
            ? 'توجه: ردیف‌های پیست‌شده تعداد ستون یکسان ندارند؛ جای ستون‌ها را بررسی کنید.' : '';
        return {
            ok: true,
            webValues: extraction.webValues,
            excelErrors: extraction.excelErrors,
            outOfRange: extraction.outOfRange,
            rowWarning,
            colWarning
        };
    }

    async function pasteFromClipboard() {
        if (state.searching || state.calculating) return;
        if (!navigator.clipboard || typeof navigator.clipboard.readText !== 'function') {
            showMainMessage('خواندن مستقیم Clipboard در این مرورگر در دسترس نیست. متن را با Ctrl+V داخل کادر پیست کنید.', 'warning');
            return;
        }
        try {
            const text = await navigator.clipboard.readText();
            if (!text) {
                showMainMessage('متن Clipboard خالی است.', 'warning');
                return;
            }
            pasteArea.value = text;
            updatePasteRevisionWarning();
            showMainMessage('متن Clipboard داخل کادر قرار گرفت؛ برای بررسی، دکمه «بررسی» را بزنید.', 'info');
        } catch (error) {
            showMainMessage('مرورگر اجازه خواندن Clipboard را نداد. متن را با Ctrl+V داخل کادر پیست کنید.', 'warning');
        }
    }

    async function reviewPastedRow() {
        if (state.searching || state.calculating) return;
        if (!activeHeaderMapping) {
            showMainMessage('ابتدا هدر اکسل را در «تنظیمات نگاشت هدر» ثبت کنید.', 'error');
            return;
        }
        const rawText = pasteArea.value;
        if (!rawText.trim()) {
            showMainMessage('کادر ورود اطلاعات خالی است. ابتدا یک ردیف از اکسل پیست کنید.', 'error');
            return;
        }

        const result = processPastedText(rawText);
        if (!result.ok) {
            showMainMessage(result.message, 'error');
            return;
        }
        if (result.colWarning) showToast(result.colWarning, 5000);

        state.webValues = result.webValues;
        state.excelErrors = result.excelErrors || {};
        state.outOfRange = result.outOfRange || {};
        state.parsed = true;
        state.reviewedText = rawText;
        state.searchConfirmed = false;
        state.headerContract = String(getVal(FIELD_KEYS.CONTRACT_TYPE) || '').trim();
        updateHeaderIdentity();
        WEB_FIELDS.forEach(field => syncControl(field.key));
        resetContractFix();
        updatePasteRevisionWarning();
        refreshAll();

        const siteIdValue = String(getVal(FIELD_KEYS.SITE_ID)).trim();
        if (!siteIdValue) {
            state.searchConfirmed = false;
            refreshAll();
            showMainMessage('مقدار Site ID در ردیف پیست‌شده وجود ندارد یا خالی است؛ جستجو انجام نشد.', 'error');
            return;
        }

        state.searching = true;
        setBusy(true);
        const search = await performAutoSearch(siteIdValue);
        state.searching = false;
        setBusy(false);

        if (!search.ok) {
            state.searchConfirmed = false;
            refreshAll();
            showMainMessage(search.reason, search.kind);
            return;
        }

        state.searchConfirmed = true;
        syncAllSelects();
        refreshAll();
        renderContractFix();
        if (result.rowWarning) {
            const identity = validateIdentity();
            showMainMessage(result.rowWarning + identity.message, identity.ok ? 'success' : 'error');
        }
    }

    // ============================================================
    // 10. APPLY CHANGES
    // ============================================================
    function dispatchFieldEvents(el) {
        const opts = { bubbles: true };
        el.dispatchEvent(new Event('input', opts));
        el.dispatchEvent(new Event('change', opts));
        el.dispatchEvent(new Event('blur', opts));
    }

    function applyOneField(field, value) {
        const info = targetInfo(field);
        const forceEditable = field.key === 'cancellationDate' && isCancellationDateEnabled();

        if (!info.exists || !info.typeMatches || (!info.editable && !forceEditable)) {
            return { applied: false, reason: 'not-editable' };
        }

        if (forceEditable && info.element && info.element.disabled) {
            info.element.disabled = false;
        }

        if (field.type === 'select') {
            const option = findSelectableOption(info.element, value);
            if (!option) return { applied: false, reason: 'invalid-select' };
            info.element.value = option.value;
            if (info.element.options[info.element.selectedIndex] !== option) {
                option.selected = true;
            }
        } else {
            info.element.value = String(value);
        }

        info.element.classList.add('epf-filled');
        return { applied: true, element: info.element, field };
    }

    function finishApply(changedCount, skippedInvalid, calcState, failedApply, calcSource) {
        closeAndReset();

        const shortList = items => items.slice(0, 6).join('، ') +
            (items.length > 6 ? ` و ${faNum(items.length - 6)} مورد دیگر` : '');
        let summary = `انتقال تمام شد: ${faNum(changedCount)} فیلد تغییر کرد.`;
        if (calcState === 'done') {
            summary += calcSource === 'form' ? ' Calculate هم با مبلغ فعلی فرم اجرا شد.' : ' Calculate هم اجرا شد.';
        }
        if (calcState === 'no-amount') summary += ' Calculate اجرا نشد: مبلغ معتبری در ردیف و فرم وجود ندارد؛ در صورت نیاز آن را دستی بزنید.';
        if (calcState === 'blocked') summary += ' Calculate اجرا نشد: درخواستی به سرور نرفت (احتمالاً اعتبارسنجی صفحه مانع شد)؛ آن را دستی بزنید.';
        if (calcState === 'error') summary += ' Calculate با خطا روبه‌رو شد؛ آن را دستی بزنید.';
        if (calcState === 'cancelled') summary += ' انتظار Calculate را لغو کردید؛ وضعیت آن را روی فرم بررسی کنید و در صورت نیاز دستی بزنید.';
        if (skippedInvalid.length) summary += ` فیلدهای لیستی با مقدار نامعتبر رد شدند: ${shortList(skippedInvalid)}.`;
        if (failedApply && failedApply.length) summary += ` فیلدهایی که هنگام اعمال خطا دادند: ${shortList(failedApply)}.`;
        summary += ' چیزی ذخیره نشده است؛ دکمهٔ Save صفحه را خودتان بزنید.';
        showToast(summary);

        pasteArea.value = '';
        resetParsedData();
        clearMainMessage();
    }

    async function seatEscalationTypeForApply(field) {
        const result = applyOneField(field, getVal(field.key));
        if (!result.applied) return { applied: false, aborted: false };

        state.cancelRequested = false;
        state.searching = true;
        setBusy(true);
        showMainMessage('در حال نشاندن Escalation Type؛ منتظر پاسخ سامانه بمانید...', 'info');
        const response = await runAndWait(() => dispatchFieldEvents(result.element), 500, 8000, () => {
            showMainMessage('سامانه هنوز پاسخ نداده است؛ ممکن است ارتباط کند یا نشست کاربری منقضی شده باشد.', 'warning');
        });
        state.searching = false;
        setBusy(false);
        syncAllSelects();

        if (response.cancelled) {
            refreshAll();
            showMainMessage('انتظار را لغو کردید؛ فقط Escalation Type نشانده شده است. وضعیت فرم را بررسی کنید و برای ادامه دوباره «بررسی» را بزنید.', 'warning');
            return { applied: true, aborted: true };
        }
        if (response.error || response.timeout) {
            refreshAll();
            showMainMessage('تغییر Escalation Type در سامانه تأیید نشد و بقیهٔ فیلدها اعمال نشدند؛ وضعیت صفحه را بررسی کنید.', 'error');
            return { applied: true, aborted: true };
        }
        return { applied: true, aborted: false, item: result };
    }

    async function applyAllFields() {
        if (state.searching || state.calculating) return;
        if (isPasteStale()) {
            showMainMessage('متن کادر پس از بررسی تغییر کرده است؛ ابتدا دوباره «بررسی» را بزنید.', 'warning');
            return;
        }
        const identity = validateIdentity();
        if (!identity.ok) {
            applyButton.disabled = true;
            showMainMessage(identity.message, 'error');
            return;
        }

        let changedCount = 0;
        const appliedItems = [];
        const skippedInvalid = [];
        const failedApply = [];
        const greenCalcKeys = [];
        let cancellationLetterPreApplied = false;

        const cancellationField = FIELD_BY_KEY.cancellationLetter;
        if (isFieldEnabled(cancellationField.key) && isCancellationDateEnabled() && evaluateField(cancellationField.key) === 'will-change') {
            try {
                const cancellationResult = applyOneField(cancellationField, getVal(cancellationField.key));
                if (cancellationResult.applied) {
                    dispatchFieldEvents(cancellationResult.element);
                    cancellationResult.alreadyDispatched = true;
                    appliedItems.push(cancellationResult);
                    changedCount += 1;
                    cancellationLetterPreApplied = true;
                }
            } catch (error) {
                failedApply.push(cancellationField.label);
            }
            syncAllSelects();
        }

        let escalationPreApplied = false;
        const escalationField = FIELD_BY_KEY[FIELD_KEYS.ESCALATION_TYPE];
        if (isFieldEnabled(escalationField.key) && evaluateField(escalationField.key) === 'will-change') {
            const pre = await seatEscalationTypeForApply(escalationField);
            if (pre.aborted) return;
            if (pre.applied) {
                escalationPreApplied = true;
                changedCount += 1;
                greenCalcKeys.push(escalationField.key);
                appliedItems.push(Object.assign({}, pre.item, { alreadyDispatched: true }));
            }
        }

        WEB_FIELDS.forEach(field => {
            if (!isFieldEnabled(field.key)) return;
            if (field.key === 'cancellationLetter' && cancellationLetterPreApplied) return;
            if (field.key === FIELD_KEYS.ESCALATION_TYPE && escalationPreApplied) return;
            const status = evaluateField(field.key);
            if (status === 'invalid-select') {
                skippedInvalid.push(field.label);
                return;
            }
            if (status !== 'will-change') return;

            if (CALC_FIELD_KEYS.includes(field.key)) greenCalcKeys.push(field.key);
            try {
                const result = applyOneField(field, getVal(field.key));
                if (result.applied) {
                    changedCount += 1;
                    appliedItems.push(result);
                }
            } catch (error) {
                failedApply.push(field.label);
            }
        });

        appliedItems.forEach(item => {
            if (item.alreadyDispatched) return;
            try {
                if (item.element && item.element.isConnected) dispatchFieldEvents(item.element);
            } catch (error) {
                failedApply.push(item.field.label);
            }
        });

        if (greenCalcKeys.length === 0) {
            finishApply(changedCount, skippedInvalid, 'none', failedApply);
            return;
        }
        const calcAmount = resolveCalcAmount();
        if (!calcAmount.valid) {
            finishApply(changedCount, skippedInvalid, 'no-amount', failedApply);
            return;
        }

        const notSatLabels = [];
        greenCalcKeys.forEach(key => {
            if (!fieldSat(FIELD_BY_KEY[key], getVal(key))) notSatLabels.push(FIELD_BY_KEY[key].label);
        });
        if (notSatLabels.length) {
            showMainMessage('خطا: مقدار این فیلدها در فرم ننشست و Calculate اجرا نشد: ' + notSatLabels.join('، '), 'error');
            refreshAll();
            return;
        }

        const calcBtn = document.getElementById(CALC_BTN_ID);
        if (!calcBtn) {
            showMainMessage('فیلدها اعمال شدند، اما دکمهٔ Calculate در صفحه پیدا نشد؛ Calculate خودکار اجرا نشد.', 'error');
            refreshAll();
            return;
        }

        state.cancelRequested = false;
        state.calculating = true;
        applyButton.disabled = true;
        setBusy(true);
        showMainMessage('در حال محاسبه (Calculate)... این مرحله چیزی ذخیره نمی‌کند.', 'info');

        runAndWait(() => calcBtn.click(), 1500, 10000, () => {
            showMainMessage('Calculate هنوز تکمیل نشده است؛ ممکن است ارتباط با سامانه کند یا قطع شده باشد.', 'warning');
        }).then(response => {
            state.calculating = false;
            setBusy(false);
            let calcState = 'done';
            if (response.cancelled) calcState = 'cancelled';
            else if (response.started === false) calcState = response.error ? 'error' : 'blocked';
            else if (response.error || response.timeout) calcState = 'error';
            finishApply(changedCount, skippedInvalid, calcState, failedApply, calcAmount.source);
        });
    }

    // ============================================================
    // 11. CONTRACT & ESCALATION HANDLING
    // ============================================================

    function buildContractBox() {
        ctBox = makeElement('div', 'epf-ct');
        ctBox.hidden = true;
        ctBtn = makeElement('button', 'epf-ct-btn');
        ctBtn.type = 'button';
        const icon = makeElement('span', '', '↻');
        icon.setAttribute('aria-hidden', 'true');
        ctBtn.append(icon, document.createTextNode('اعمال Contract جدید'));
        ctMsg = makeElement('div', 'epf-ct-msg');
        ctMsg.setAttribute('role', 'status');
        ctBtn.addEventListener('click', applyContractChange);
        ctBox.append(ctBtn, ctMsg);
        return ctBox;
    }

    function setCtMsg(text, kind) {
        if (!ctMsg) return;
        ctMsg.textContent = text || '';
        ctMsg.className = 'epf-ct-msg' + (kind ? ' epf-ct-' + kind : '');
    }

    function resetContractFix() {
        if (!ctBox) return;
        ctBox.hidden = true;
        ctBtn.hidden = false;
        ctBtn.disabled = false;
        setCtMsg('');
    }

    function renderContractFix() {
        if (!ctBox) return;
        if (state.contractChangeCompleted) {
            ctBox.hidden = true;
            return;
        }
        resetContractFix();
        if (!state.parsed) return;
        const siteInfo = targetInfo(FIELD_BY_KEY[FIELD_KEYS.SITE_ID]);
        const ctInfo = targetInfo(FIELD_BY_KEY[FIELD_KEYS.CONTRACT_TYPE]);
        if (!siteInfo.exists || !siteInfo.typeMatches || !ctInfo.exists || !ctInfo.typeMatches) return;

        const pastedSite = normalize(getVal(FIELD_KEYS.SITE_ID));
        if (!pastedSite || pastedSite !== normalize(siteInfo.element.value)) return;
        const pastedCt = getVal(FIELD_KEYS.CONTRACT_TYPE);
        if (normalize(pastedCt) === '') return;
        if (normalize(pastedCt) === normalize(selectedOptionText(ctInfo.element))) return;

        ctBox.hidden = false;
        const opt = findOptionByText(ctInfo.element, pastedCt);
        if (!opt) {
            ctBtn.disabled = true;
            setCtMsg('در لیست سایت نیست', 'bad');
            return;
        }
        if (DATE_OPTIONAL_TYPES.indexOf(String(opt.textContent).trim().toUpperCase()) !== -1) {
            setCtMsg('با این تغییر، تاریخ‌ها غیرفعال و پاک می‌شوند', 'info');
        }
    }

    async function applyContractChange() {
        const info = targetInfo(FIELD_BY_KEY[FIELD_KEYS.CONTRACT_TYPE]);
        if (!info.exists || !info.typeMatches) return;
        const opt = findOptionByText(info.element, getVal(FIELD_KEYS.CONTRACT_TYPE));
        if (!opt) {
            ctBtn.disabled = true;
            setCtMsg('در لیست سایت نیست', 'bad');
            return;
        }
        if (!info.editable) {
            setCtMsg('این فیلد در فرم قابل ویرایش نیست', 'bad');
            return;
        }
        const el = info.element;
        try {
            el.value = opt.value;
            if (el.options[el.selectedIndex] !== opt) opt.selected = true;
            el.classList.add('epf-filled');
            dispatchFieldEvents(el);
        } catch (error) {
            setCtMsg('تغییر نوع قرارداد در سامانه تأیید نشد. وضعیت صفحه را بررسی کنید و در صورت نیاز دوباره تلاش کنید.', 'bad');
            return;
        }

        ctBtn.hidden = true;
        setCtMsg('✓ نوع قرارداد تغییر کرد', 'ok');
        markContractChanged(String(opt.textContent || '').trim());
        syncAllSelects();
        refreshAll();
        await seatEscalationTypeAfterContract();
    }

    function markContractChanged(contractText) {
        state.contractChangeCompleted = true;
        state.headerContract = contractText || '';
        updateHeaderIdentity();
        if (contractFixTimer) clearTimeout(contractFixTimer);
        contractFixTimer = setTimeout(() => {
            state.contractChangeCompleted = false;
            contractFixTimer = null;
        }, 8000);
    }

    async function seatEscalationTypeAfterContract() {
        if (!state.parsed) return;
        const pasted = getVal(FIELD_KEYS.ESCALATION_TYPE);
        if (normalize(pasted) === '') return;

        const field = FIELD_BY_KEY[FIELD_KEYS.ESCALATION_TYPE];
        if (!isFieldEnabled(field.key)) return;
        const info = targetInfo(field);
        if (!info.exists || !info.typeMatches || !info.editable) return;
        if (fieldSat(field, pasted)) return;

        const option = findSelectableOption(info.element, pasted);
        if (!option) {
            setCtMsg('Escalation Type پیست‌شده در لیست صفحه نیست؛ هنگام «اعمال در فرم وب» بررسی می‌شود.', 'info');
            return;
        }

        setCtMsg('در حال نشاندن Escalation Type در فرم...', 'info');
        const esc = info.element;
        esc.value = option.value;
        if (esc.options[esc.selectedIndex] !== option) option.selected = true;
        esc.classList.add('epf-filled');
        state.cancelRequested = false;
        state.searching = true;
        setBusy(true);
        const result = await runAndWait(() => dispatchFieldEvents(esc), 500, 8000, () => {
            setCtMsg('سامانه هنوز پاسخ نداده است؛ ممکن است نشست کاربری منقضی یا ارتباط قطع شده باشد.', 'info');
        });
        state.searching = false;
        setBusy(false);
        if (result.cancelled) {
            setCtMsg('انتظار را لغو کردید؛ وضعیت Escalation Type و باکس Escalation % را روی فرم بررسی کنید.', 'info');
            syncAllSelects();
            refreshAll();
            return;
        }
        if (result.error || result.timeout) {
            setCtMsg('تغییر Escalation Type در سامانه تأیید نشد. وضعیت صفحه را بررسی کنید.', 'bad');
            return;
        }

        if (fieldSat(field, pasted)) {
            const fresh = document.getElementById(field.id);
            if (fresh) fresh.classList.add('epf-filled');
        }
        syncAllSelects();
        refreshAll();
        const percentInfo = targetInfo(FIELD_BY_KEY.escalation);
        const percentOpen = percentInfo.exists && percentInfo.editable;
        setCtMsg(percentOpen
            ? '✓ نوع قرارداد و Escalation Type تغییر کردند. باکس Escalation % روی فرم باز شد؛ اکنون «اعمال در فرم وب» را بزنید.'
            : '✓ نوع قرارداد و Escalation Type تغییر کردند؛ اکنون «اعمال در فرم وب» را بزنید.', 'ok');
    }

    // ============================================================
    // 12. HEADER MAPPING SETTINGS & FIELD TOGGLES
    // ============================================================
    const settingsUi = {
        built: false,
        dirty: false,
        defaultDraft: false,
        lastResult: null,
        checkedText: null,
        headerArea: null,
        statusBox: null,
        grid: null,
        dupLegendEl: null,
        saveButton: null,
        fieldCheckboxes: {}
    };

    function setSettingsStatus(text, kind) {
        const box = settingsUi.statusBox;
        if (!box) return;
        box.textContent = text || '';
        box.className = text ? 'epf-box epf-show epf-box-' + (kind || 'info') : 'epf-box';
    }

    function updateDupLegendCount(dupCount) {
        if (!settingsUi.dupLegendEl) return;
        const countStr = Number.isInteger(dupCount) && dupCount > 0
            ? ` (${faNum(dupCount)} مورد؛ اولین ستون استفاده می‌شود)` : ' (اولین ستون استفاده می‌شود)';
        settingsUi.dupLegendEl.textContent = 'نارنجی: هدر تکراری' + countStr;
    }

    function renderHeaderGrid(result) {
        if (!settingsUi.grid || !settingsUi.siteRow) return;
        settingsUi.siteRow.textContent = '';
        settingsUi.grid.textContent = '';
        const dupCount = result ? (result.dupCount || 0) : 0;
        updateDupLegendCount(dupCount);

        WEB_FIELDS.forEach((field, index) => {
            const isEnabled = isFieldEnabled(field.key);
            const st = result && result.states ? result.states[field.key] : { st: isEnabled ? 'ok' : 'disabled', col: isEnabled ? index : null };
            const cell = makeElement('div', 'epf-hcell');
            cell.title = field.label;

            const leftBox = makeElement('div', 'epf-hcell-left');
            const chk = makeElement('input', 'epf-hchk');
            chk.type = 'checkbox';
            chk.checked = isEnabled;
            chk.dataset.key = field.key;
            if (field.key === FIELD_KEYS.SITE_ID || field.key === FIELD_KEYS.CONTRACT_TYPE) {
                chk.disabled = true;
                chk.title = 'فیلد الزامی هویتی (غیرقابل تغییر)';
            }
            chk.addEventListener('change', () => {
                if (chk.checked) activeDisabledFields.delete(field.key);
                else activeDisabledFields.add(field.key);
                onHeaderDraftChanged();
            });
            settingsUi.fieldCheckboxes[field.key] = chk;

            const name = makeElement('span', 'epf-hname', field.label);
            const col = makeElement('span', 'epf-hcol');

            if (!isEnabled || st.st === 'disabled') {
                name.classList.add('epf-h-disabled');
                col.textContent = '✕';
                col.classList.add('epf-h-disabled');
            } else if (st.st === 'missing') {
                const missingClass = REQUIRED_HEADER_KEYS.includes(field.key) ? 'epf-h-missing' : 'epf-h-absent';
                name.classList.add(missingClass);
                col.textContent = '—';
                col.classList.add(missingClass);
            } else if (st.st === 'dup') {
                name.classList.add('epf-h-dup');
                col.textContent = Number.isInteger(st.col) ? String(st.col + 1) : '—';
                col.classList.add('epf-h-dup');
            } else {
                col.textContent = Number.isInteger(st.col) ? String(st.col + 1) : '—';
                if (st.col !== index) col.classList.add('epf-h-changed');
            }

            leftBox.append(chk, name);
            cell.append(leftBox, col);
            if (field.key === FIELD_KEYS.SITE_ID) {
                settingsUi.siteRow.appendChild(cell);
            } else {
                settingsUi.grid.appendChild(cell);
            }
        });
    }

    function onHeaderDraftChanged() {
        settingsUi.dirty = true;
        settingsUi.defaultDraft = false;
        settingsUi.lastResult = null;
        settingsUi.checkedText = null;
        settingsUi.saveButton.disabled = true;
        renderHeaderGrid(null);
        setSettingsStatus('تنظیمات یا متن هدر تغییر کرده است؛ برای بررسی، «بررسی هدر» را بزنید.', 'info');
    }

    function checkHeaderDraft() {
        const raw = settingsUi.headerArea.value;
        if (!raw.trim()) {
            settingsUi.defaultDraft = false;
            settingsUi.lastResult = null;
            settingsUi.checkedText = null;
            settingsUi.saveButton.disabled = true;
            renderHeaderGrid(null);
            setSettingsStatus('کادر هدر خالی است؛ ردیف عنوان ستون‌ها را از اکسل پیست کنید.', 'warn');
            return;
        }

        const result = buildHeaderMapping(raw);
        settingsUi.lastResult = result;
        settingsUi.checkedText = raw;
        renderHeaderGrid(result.parsed ? result : null);

        if (!result.parsed) {
            settingsUi.saveButton.disabled = true;
            setSettingsStatus('فقط یک ردیف هدر باید ثبت شود؛ ردیف‌های اضافی را حذف کنید.', 'err');
            return;
        }
        if (!result.ready) {
            settingsUi.saveButton.disabled = true;
            const missingReq = [];
            if (!result.hasSiteId) missingReq.push('Site ID');
            if (!result.hasContractType) missingReq.push('Contract Type');
            if (missingReq.length) {
                setSettingsStatus('هدر الزامی یافت نشد: ' + missingReq.join(' و ') + '؛ وجود هر دو هدر Site ID و Contract Type به‌همراه دست‌کم یک هدر دیگر الزامی است.', 'err');
            } else {
                setSettingsStatus('حداقل ۳ هدر شناخته‌شده فعال لازم است (Site ID، Contract Type و دست‌کم یک هدر دیگر)؛ فعلاً ' + faNum(result.mappedCount) + ' هدر شناسایی شد.', 'err');
            }
            return;
        }
        const dupCount = result.dupCount || 0;
        settingsUi.saveButton.disabled = false;
        let summary = faNum(result.mappedCount) + ' فیلد از ' + faNum(result.enabledCount) + ' فیلد فعال شناسایی شد';
        if (result.missing > 0) summary += `؛ ${faNum(result.missing)} فیلد فعال بدون هدر (زرد) بررسی نمی‌شوند`;
        if (dupCount) {
            setSettingsStatus(summary + `؛ ${faNum(dupCount)} هدر تکراری است (نارنجی) و اولین ستون آن استفاده می‌شود. برای ثبت، «ذخیره هدر» را بزنید.`, 'warn');
        } else {
            setSettingsStatus(summary + '؛ برای ثبت، «ذخیره هدر» را بزنید.', 'ok');
        }
    }

    function resetHeaderDraft() {
        activeDisabledFields.clear();
        settingsUi.headerArea.value = '';
        settingsUi.dirty = true;
        settingsUi.defaultDraft = true;
        settingsUi.lastResult = defaultMapping();
        settingsUi.checkedText = '';
        settingsUi.saveButton.disabled = false;
        renderHeaderGrid(settingsUi.lastResult);
        setSettingsStatus('حالت پیش‌فرض (همه ۳۶ فیلد فعال، ستون‌های ۱ تا ۳۶) آماده است؛ برای فعال‌سازی، «ذخیره هدر» را بزنید.', 'info');
    }

    async function pasteHeaderFromClipboard() {
        if (!navigator.clipboard || typeof navigator.clipboard.readText !== 'function') {
            setSettingsStatus('خواندن مستقیم Clipboard در این مرورگر در دسترس نیست؛ با Ctrl+V داخل کادر پیست کنید.', 'warn');
            return;
        }
        try {
            const text = await navigator.clipboard.readText();
            if (!text) {
                setSettingsStatus('متن Clipboard خالی است.', 'warn');
                return;
            }
            settingsUi.headerArea.value = text;
            onHeaderDraftChanged();
            setSettingsStatus('متن Clipboard داخل کادر قرار گرفت؛ برای بررسی، «بررسی هدر» را بزنید.', 'info');
        } catch (error) {
            setSettingsStatus('مرورگر اجازه خواندن Clipboard را نداد؛ با Ctrl+V داخل کادر پیست کنید.', 'warn');
        }
    }

    function buildSettings() {
        if (settingsUi.built) return;
        settingsUi.built = true;

        const intro = makeElement('p', 'epf-settings-intro',
            'ردیف عنوان ستون‌های اکسل (Header) را کپی کنید، با دکمهٔ «پیست» در کادر زیر بریزید و «بررسی هدر» را بزنید. ' +
            'هدر Site ID و Contract Type به‌همراه دست‌کم یک هدر دیگر الزامی است (حداقل ۳ هدر). ' +
            'تیکِ فیلدهایی که نیازی به آن‌ها ندارید را بردارید تا کلاً در پنل مخفی شده و بررسی نشوند. ' +
            'تطبیق نام‌ها به حروف بزرگ و کوچک و فاصلهٔ اضافی حساس نیست.');

        const label = makeElement('label', 'epf-paste-label', 'ردیف هدر اکسل');
        label.htmlFor = 'epf-header-area';
        const area = makeElement('textarea', 'epf-settings-area');
        area.id = 'epf-header-area';
        area.spellcheck = false;
        area.placeholder = 'Site ID\tContract Type\tLL Name\tOwner Type\t...';
        area.setAttribute('aria-label', 'ردیف هدر اکسل');
        settingsUi.headerArea = area;
        area.addEventListener('input', onHeaderDraftChanged);

        const layout = makeElement('div', 'epf-settings-layout');
        const actions = makeElement('div', 'epf-settings-actions');
        const pasteBtn = makeElement('button', 'epf-btn epf-btn-primary', 'پیست');
        pasteBtn.type = 'button';
        pasteBtn.addEventListener('click', pasteHeaderFromClipboard);
        const checkBtn = makeElement('button', 'epf-btn epf-btn-primary', 'بررسی هدر');
        checkBtn.type = 'button';
        checkBtn.addEventListener('click', checkHeaderDraft);
        actions.append(pasteBtn, checkBtn);
        layout.append(area, actions);

        settingsUi.statusBox = makeElement('div', 'epf-box');
        settingsUi.statusBox.setAttribute('role', 'status');

        const legend = makeElement('div', 'epf-hlegend');
        legend.appendChild(makeElement('span', 'lg-title', 'راهنمای جدول:'));
        legend.appendChild(makeElement('span', 'lg-def', 'خاکستری: ستون پیش‌فرض'));
        legend.appendChild(makeElement('span', 'lg-chg', 'سبز: شمارهٔ ستون تغییر کرده'));
        legend.appendChild(makeElement('span', 'lg-abs', 'زرد: بدون هدر؛ بررسی نمی‌شود'));
        settingsUi.dupLegendEl = makeElement('span', 'lg-dup', 'نارنجی: هدر تکراری (اولین ستون استفاده می‌شود)');
        legend.appendChild(settingsUi.dupLegendEl);
        legend.appendChild(makeElement('span', 'lg-dis', 'خط‌خورده: غیرفعال شده توسط کاربر'));
        legend.appendChild(makeElement('span', 'lg-miss', 'قرمز: هدر الزامی یافت نشد'));

        const gridWrap = makeElement('div', 'epf-hgrid-wrap');
        settingsUi.siteRow = makeElement('div', 'epf-hgrid-site-row');
        settingsUi.grid = makeElement('div', 'epf-hgrid');
        gridWrap.append(settingsUi.siteRow, settingsUi.grid);

        settingsView.append(intro, label, layout, settingsUi.statusBox, legend, gridWrap);

        const resetButton = makeElement('button', 'epf-btn epf-btn-reset', 'بازگشت به پیش‌فرض');
        resetButton.type = 'button';
        resetButton.addEventListener('click', resetHeaderDraft);

        const saveButton = makeElement('button', 'epf-btn epf-btn-apply', 'ذخیره هدر');
        saveButton.type = 'button';
        saveButton.addEventListener('click', saveHeaderSettings);
        const cancelButton2 = makeElement('button', 'epf-btn epf-btn-cancel', 'انصراف');
        cancelButton2.type = 'button';
        cancelButton2.addEventListener('click', leaveSettings);
        settingsUi.saveButton = saveButton;
        settingsFooter.append(resetButton, cancelButton2, saveButton);
    }

    function updateGearState() {
        let mode = 'none';
        let text = 'ثبت نشده';
        if (activeHeaderMapping) {
            mode = isDefaultMapping(activeHeaderMapping) ? 'default' : 'custom';
            text = mode === 'default' ? 'پیش‌فرض' : 'سفارشی شده';
        }
        gearBtn.dataset.mode = mode;
        gearState.textContent = text;
    }

    function updateMapNote() {
        const lines = [];
        if (initNote) lines.push(initNote);
        if (!activeHeaderMapping && !initNote) lines.push('برای شروع، ابتدا ردیف هدر اکسل را از «تنظیمات نگاشت هدر» ثبت کنید.');
        mapNote.textContent = lines.join('\n');
        mapNote.classList.toggle('epf-show', lines.length > 0);
        reviewButton.disabled = !activeHeaderMapping || state.searching || state.calculating;
    }

    function openSettings() {
        if (state.searching || state.calculating || state.settingsOpen) return;
        buildSettings();
        settingsUi.headerArea.value = activeHeaderText;
        settingsUi.dirty = false;
        settingsUi.defaultDraft = false;

        if (activeHeaderMapping) {
            const result = buildHeaderMapping(activeHeaderText);
            settingsUi.lastResult = result;
            settingsUi.checkedText = activeHeaderText;
            renderHeaderGrid(result);
            setSettingsStatus('هدر ثبت‌شده فعلی نمایش داده می‌شود؛ برای تغییر، هدر یا چک‌باکس‌های فیلدها را تغییر دهید.', 'info');
            settingsUi.saveButton.disabled = false;
        } else {
            settingsUi.lastResult = null;
            settingsUi.checkedText = null;
            settingsUi.saveButton.disabled = true;
            renderHeaderGrid(null);
            setSettingsStatus('هدر ثبت نشده است؛ ردیف عنوان ستون‌ها را پیست و «بررسی هدر» را بزنید.', 'info');
        }

        state.settingsOpen = true;
        headerMain.hidden = true;
        gearBtn.hidden = true;
        closeMainBtn.hidden = true;
        closeSettingsBtn.hidden = false;
        headerIdentity.hidden = true;
        content.hidden = true;
        footer.hidden = true;
        settingsView.hidden = false;
        settingsFooter.hidden = false;
        title.textContent = 'تنظیمات نگاشت هدر';
        settingsUi.headerArea.focus();
    }

    function hideSettings() {
        state.settingsOpen = false;
        settingsView.hidden = true;
        settingsFooter.hidden = true;
        headerMain.hidden = false;
        gearBtn.hidden = false;
        closeMainBtn.hidden = false;
        closeSettingsBtn.hidden = true;
        headerIdentity.hidden = false;
        content.hidden = false;
        footer.hidden = false;
        title.textContent = 'انتقال اطلاعات Excel به فرم BTS Property';
        updateMapNote();
    }

    async function leaveSettings() {
        if (!state.settingsOpen || state.confirmOpen) return;
        if (settingsUi.dirty) {
            const leave = await askConfirm('تغییرات ذخیره‌نشده حذف می‌شود. به پنل اصلی برگردیم؟', 'برگشت بدون ذخیره', 'ماندن در تنظیمات');
            if (!leave) return;
        }
        activeDisabledFields = loadStoredDisabledFields();
        hideSettings();
        refreshAll();
        gearBtn.focus();
    }

    function saveHeaderSettings() {
        let text = settingsUi.headerArea.value;
        let result;
        if (settingsUi.defaultDraft && !text.trim()) {
            text = WEB_FIELDS.map(field => field.label).join('\t');
            result = defaultMapping();
        } else {
            if (!settingsUi.lastResult || !settingsUi.lastResult.ready || settingsUi.checkedText !== text) return;
            result = settingsUi.lastResult;
        }

        persistDisabledFields(activeDisabledFields);

        activeHeaderText = text;
        activeHeaderMapping = { columns: result.columns };
        const stored = persistHeader(activeHeaderText);
        initNote = stored ? '' : 'ذخیره دائمی هدر ممکن نیست؛ هدر فقط تا پایان این نشست فعال است.';
        updateGearState();

        state.webValues = {};
        state.parsed = false;
        state.reviewedText = '';
        state.searchConfirmed = false;
        settingsUi.dirty = false;
        hideSettings();
        resetParsedData();
        clearMainMessage();
        updateMapNote();
        reviewButton.disabled = false;
        pasteArea.focus();
        const savedCount = result.mappedCount || result.enabledCount;
        showMainMessage('هدر و تنظیمات فیلدها ذخیره شد؛ ' + faNum(savedCount) + ' فیلد از ' + faNum(result.enabledCount) + ' فیلد فعال نگاشت شد' +
            (stored ? ' و در حافظهٔ مرورگر ماند.' : '؛ ذخیرهٔ دائمی ممکن نیست و تنظیمات فقط تا پایان این نشست فعال است.'), stored ? 'success' : 'warning');
    }

    // ============================================================
    // 13. PANEL LIFECYCLE (inline)
    // ============================================================
    function showPanel() {
        if (state.isOpen) return;
        if (state.settingsOpen) hideSettings();
        state.isOpen = true;
        overlay.classList.remove('epf-closing');
        overlay.classList.add('epf-open');
        overlay.setAttribute('aria-hidden', 'false');
        if (headerBtn) headerBtn.setAttribute('aria-expanded', 'true');

        if (state.parsed) {
            const siteInfo = targetInfo(FIELD_BY_KEY[FIELD_KEYS.SITE_ID]);
            const pastedSite = normalize(getVal(FIELD_KEYS.SITE_ID));
            const currentSite = siteInfo.exists ? normalize(siteInfo.element.value) : '';
            if (pastedSite !== currentSite) {
                state.searchConfirmed = false;
            }
            syncAllSelects();
            updateHeaderIdentity();
            refreshAll();
            renderContractFix();
        }

        requestAnimationFrame(() => {
            const offsetPx = 64 + 56 + 20;
            try {
                overlay.scrollIntoView({ behavior: 'smooth', block: 'start' });
            } catch (e) {
                try {
                    window.scrollTo({ top: overlay.offsetTop - offsetPx, behavior: 'smooth' });
                } catch (e2) { }
            }
        });
        setTimeout(() => { try { pasteArea.focus({ preventScroll: true }); } catch (e) { } }, 80);
    }

    function hidePanel(animate) {
        if (!state.isOpen) return;
        state.isOpen = false;
        overlay.setAttribute('aria-hidden', 'true');
        if (headerBtn) headerBtn.setAttribute('aria-expanded', 'false');
        if (animate === false) {
            overlay.classList.remove('epf-open', 'epf-closing');
            return;
        }
        overlay.classList.add('epf-closing');
        setTimeout(() => {
            if (!state.isOpen) {
                overlay.classList.remove('epf-open', 'epf-closing');
            }
        }, 200);
    }

    function closeAndReset() {
        if (!state.isOpen) return;
        if (state.searching || state.calculating) {
            state.searching = false;
            state.calculating = false;
            state.cancelRequested = false;
        }
        if (state.settingsOpen) hideSettings();
        hidePanel();
        pasteArea.value = '';
        resetParsedData();
        clearMainMessage();
        if (headerBtn) headerBtn.focus();
    }

    function togglePanel() {
        if (state.isOpen) {
            hidePanel();
        } else {
            showPanel();
        }
    }

    function scrollToSaveButton() {
        const saveBtn = document.getElementById('Content_btn_Save');
        if (!saveBtn) {
            showToast('دکمه ذخیره در این صفحه پیدا نشد.', 4000);
            return;
        }
        try {
            saveBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
        } catch (e) {
            const offsetPx = window.innerHeight / 2;
            try {
                window.scrollTo({ top: saveBtn.offsetTop - offsetPx, behavior: 'smooth' });
            } catch (e2) { }
        }
        const prevOutline = saveBtn.style.outline;
        const prevOutlineOffset = saveBtn.style.outlineOffset;
        saveBtn.style.outline = '3px solid #ff8c00';
        saveBtn.style.outlineOffset = '2px';
        setTimeout(() => {
            saveBtn.style.outline = prevOutline;
            saveBtn.style.outlineOffset = prevOutlineOffset;
        }, 1500);
    }

    function waitUntilIdle(maxMs) {
        return new Promise(resolve => {
            const start = Date.now();
            const check = () => {
                if ((!state.searching && !state.calculating) || Date.now() - start >= maxMs) { resolve(); return; }
                setTimeout(check, 40);
            };
            check();
        });
    }

    async function requestCancelAndClose() {
        if (state.confirmOpen) return;
        if (state.searching || state.calculating) {
            const leave = await askConfirm('عملیات هنوز در انتظار پاسخ سامانه است. انتظار لغو شود و پنل بسته شود؟ وضعیت فرم را بعداً بررسی کنید.', 'لغو و بستن', 'ادامه انتظار');
            if (!leave) return;
            cancelActiveWait();
            await waitUntilIdle(1500);
            closeAndReset();
            return;
        }
        if (state.parsed) {
            const ok = await askConfirm('اطلاعات بررسی‌شده پاک می‌شود. پنل بسته شود؟', 'بستن پنل', 'ماندن');
            if (!ok) return;
        }
        closeAndReset();
    }

    let toastTimer = null;
    function showToast(message, duration) {
        toast.textContent = message;
        toast.classList.add('epf-toast-show');
        if (toastTimer) clearTimeout(toastTimer);
        toastTimer = setTimeout(() => toast.classList.remove('epf-toast-show'), duration || 8000);
    }

    function askConfirm(message, okLabel, stayLabel) {
        return new Promise(resolve => {
            confirmMsg.textContent = message;
            confirmOkBtn.textContent = okLabel;
            confirmStayBtn.textContent = stayLabel;
            confirmOverlay.classList.add('epf-open');
            confirmOverlay.setAttribute('aria-hidden', 'false');
            state.confirmOpen = true;

            const finish = result => {
                confirmOverlay.classList.remove('epf-open');
                confirmOverlay.setAttribute('aria-hidden', 'true');
                state.confirmOpen = false;
                confirmOkBtn.removeEventListener('click', onOk);
                confirmStayBtn.removeEventListener('click', onStay);
                confirmOverlay.removeEventListener('click', onBackdrop);
                document.removeEventListener('keydown', onKey, true);
                resolve(result);
            };
            const onOk = () => finish(true);
            const onStay = () => finish(false);
            const onBackdrop = event => { if (event.target === confirmOverlay) finish(false); };
            const onKey = event => {
                if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); finish(false); }
            };

            confirmOkBtn.addEventListener('click', onOk);
            confirmStayBtn.addEventListener('click', onStay);
            confirmOverlay.addEventListener('click', onBackdrop);
            document.addEventListener('keydown', onKey, true);
            confirmStayBtn.focus();
        });
    }

    function updateHeaderIdentity() {
        const site = String(getVal(FIELD_KEYS.SITE_ID) || '').trim();
        const contract = String(state.headerContract || getVal(FIELD_KEYS.CONTRACT_TYPE) || '').trim();
        headerIdentity.textContent = (site || contract) ? (site || '—') + ' - ' + (contract || '—') : '— —';
    }

    // ============================================================
    // 14. PAGE UPDATE HANDLING
    // ============================================================
    function scheduleRefresh() {
        if (!state.isOpen || state.refreshQueued || state.searching || state.calculating) return;
        state.refreshQueued = true;
        setTimeout(() => {
            state.refreshQueued = false;
            if (!state.isOpen) return;
            if (state.parsed) syncAllSelects();
            updateHeaderIdentity();
            refreshAll();
            renderContractFix();
        }, 40);
    }

    const identityIds = new Set([
        FIELD_BY_KEY[FIELD_KEYS.SITE_ID].id,
        FIELD_BY_KEY[FIELD_KEYS.CONTRACT_TYPE].id,
        LL_NAME_ID
    ]);

    function onIdentityFieldChanged(event) {
        if (event.target && identityIds.has(event.target.id)) {
            if (!state.searching && event.target.id === SITE_ID_INPUT_ID) state.searchConfirmed = false;
            scheduleRefresh();
        }
    }

    document.addEventListener('input', onIdentityFieldChanged, true);
    document.addEventListener('change', onIdentityFieldChanged, true);

    function hookPageRefresh() {
        const prm = getPRM();
        if (prm && typeof prm.add_endRequest === 'function') {
            prm.add_endRequest(scheduleRefresh);
            return;
        }
        const panel = document.getElementById('Content_UpdatePanelInsert');
        if (panel) new MutationObserver(scheduleRefresh).observe(panel, { childList: true, subtree: true });
    }

    // ============================================================
    // 15. EVENT HANDLERS
    // ============================================================
    closeMainBtn.addEventListener('click', () => { hidePanel(); });
    closeSettingsBtn.addEventListener('click', () => { leaveSettings(); });
    pasteButton.addEventListener('click', pasteFromClipboard);
    reviewButton.addEventListener('click', reviewPastedRow);
    pasteArea.addEventListener('input', updatePasteRevisionWarning);
    applyButton.addEventListener('click', () => {
        applyAllFields().catch(error => {
            state.searching = false;
            state.calculating = false;
            setBusy(false);
            showMainMessage('خطای پیش‌بینی‌نشده هنگام اعمال اطلاعات رخ داد. وضعیت فرم را بررسی کنید.', 'error');
        });
    });
    gotoSaveButton.addEventListener('click', scrollToSaveButton);
    cancelButton.addEventListener('click', () => {
        if (state.searching || state.calculating) cancelActiveWait();
        else requestCancelAndClose();
    });
    onlyBox.addEventListener('change', updateGroups);
    gearBtn.addEventListener('click', openSettings);
    toast.addEventListener('click', () => toast.classList.remove('epf-toast-show'));

    modal.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' && event.target && (event.target.tagName === 'INPUT' || event.target.tagName === 'SELECT')) {
            event.preventDefault();
            event.stopPropagation();
        }
    });

    document.addEventListener('keydown', event => {
        if (event.key !== 'Escape' || !state.isOpen || state.confirmOpen) return;
        if (state.settingsOpen) { leaveSettings(); return; }
    });

    // ============================================================
    // 16. INJECTION & INITIALIZATION
    // ============================================================

    function injectHeaderButton() {
        if (document.getElementById('epfMenuLink')) {
            headerBtn = document.getElementById('epf-header-btn');
            return true;
        }
        const nav = document.getElementById('topmenuNav');
        if (!nav) return false;

        const reportLink = nav.querySelector('#reportLink');
        const spacer = nav.querySelector('.topmenu-spacer');

        const li = document.createElement('li');
        li.id = 'epfMenuLink';

        const btn = document.createElement('button');
        btn.id = 'epf-header-btn';
        btn.type = 'button';
        btn.className = 'nav-link epf-nav-link';
        btn.setAttribute('aria-controls', 'epf-modal');
        btn.setAttribute('aria-expanded', 'false');
        const icon = makeElement('i', 'fas fa-file-excel');
        icon.setAttribute('aria-hidden', 'true');
        const label = makeElement('span', 'link-text', 'انتقال از Excel');
        btn.append(icon, label);
        btn.addEventListener('click', togglePanel);

        li.appendChild(btn);

        if (spacer) {
            nav.insertBefore(li, spacer);
        } else if (reportLink) {
            reportLink.insertAdjacentElement('afterend', li);
        } else {
            nav.appendChild(li);
        }

        headerBtn = btn;
        return true;
    }

    function injectPanelRoot() {
        if (root.parentNode) return true;
        const card = document.querySelector('.main-card');
        if (!card) return false;
        const headerBar = card.querySelector('.card-header-custom');
        const tabsWrapper = card.querySelector('.tabs-wrapper');

        if (headerBar && tabsWrapper) {
            card.insertBefore(root, tabsWrapper);
            return true;
        }
        if (headerBar) {
            headerBar.insertAdjacentElement('afterend', root);
            return true;
        }
        return false;
    }

    function tryInit() {
        let attempts = 0;
        const maxAttempts = 100;
        const attempt = () => {
            const rootOk = injectPanelRoot();
            const btnOk = injectHeaderButton();
            if (rootOk && btnOk) {
                if (!built) buildRows();
                hookPageRefresh();
                updateMapNote();
                updateGearState();
                updateChips(null);
                return;
            }
            if (++attempts >= maxAttempts) {
                console.warn('[EPF] Could not locate .main-card / .card-header-custom after 20s.');
                return;
            }
            setTimeout(attempt, 200);
        };
        attempt();
    }

    (document.head || document.documentElement).appendChild(styleElement);
    tryInit();
})();
