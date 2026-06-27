// ==UserScript==
// @name         畅言加好友 阿陌专用 后台稳定版
// @namespace    http://tampermonkey.net/
// @version      10.1.3
// @description  畅言加好友阿陌专用，两弹窗慢等+话术确认后再完成
// @match        *://web.rvtqh.com/*
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @grant        none
// @run-at       document-end
// @homepageURL  https://github.com/a18279023705-cmd/changyan-update
// @updateURL    https://raw.githubusercontent.com/a18279023705-cmd/changyan-update/main/userscript/changyan-add-friend.meta.js
// @downloadURL  https://github.com/a18279023705-cmd/changyan-update/releases/latest/download/changyan-add-friend.user.js
// ==/UserScript==

(function () {
    'use strict';

    function readScriptVersionFromMeta() {
        try {
            if (typeof GM_info !== 'undefined' && GM_info?.script?.version) {
                return String(GM_info.script.version).trim();
            }
        } catch (e) { /* ignore */ }
        try {
            const scripts = document.getElementsByTagName('script');
            for (let i = scripts.length - 1; i >= 0; i--) {
                const text = scripts[i].textContent || '';
                if (!text.includes('changyan_add_friend') && !text.includes('畅言加好友')) continue;
                const m = text.match(/@version\s+([^\s\n\r]+)/);
                if (m) return m[1].trim();
            }
        } catch (e) { /* ignore */ }
        return '0.0.0';
    }

    const SCRIPT_VERSION = readScriptVersionFromMeta();
    const RAW_BASE =
        'https://raw.githubusercontent.com/a18279023705-cmd/changyan-update/main/userscript';
    const CDN_BASE =
        'https://cdn.jsdelivr.net/gh/a18279023705-cmd/changyan-update@main/userscript';
    const VERSION_URL = CDN_BASE + '/changyan-add-friend.version.txt';
    const MIN_VERSION_URL = CDN_BASE + '/changyan-add-friend.min-version.txt';
    const DOWNLOAD_URL =
        'https://github.com/a18279023705-cmd/changyan-update/releases/latest/download/changyan-add-friend.user.js';

    /** 操作节奏（无频繁随机等待） */
    const PACE = {
        stuckForceMs: 18000,
        afterSuccessMs: 900,
        afterNotFoundMs: 550,
        afterSkipMs: 400,
        afterRetryMs: 900,
        searchWaitMs: 520,
        searchProfileWaitMs: 12000,
        actionWaitMs: 900,
        friendApplyWaitMs: 8000,
        actionSettleMs: 500,
        /** 弹窗1：搜索结果/资料页出现后再点「添加好友」 */
        profilePanelSettleMs: 2000,
        /** 弹窗1→2：点「添加好友」后等申请页弹出 */
        afterAddFriendClickMs: 2500,
        /** 弹窗2：申请页完全打开后再填话术、点完成 */
        friendApplySettleMs: 2800,
        confirmWaitMs: 800,
        submitSettleMinMs: 1200,
        submitWaitMs: 18000,
        confirmSettleMs: 450,
        panelCloseMs: 600,
        panelCloseMinMs: 450,
        panelCloseRetryMs: 550,
        remarkCheckMs: 500,
        remarkSettleMs: 1200,
        remarkFillWaitMs: 12000,
        remarkBeforeCompleteMs: 2000,
        remarkFillRetries: 3,
        confirmClickMs: 600,
        typingMs: 120,
        typingAfterMs: 200,
        pollFastMs: 250,
        pollNormalMs: 320,
    };

    const STORAGE_KEY = 'changyan_add_friend_talks';
    const PROGRESS_STORAGE_KEY = 'changyan_add_friend_progress';
    const PHONE_LIST_STORAGE_KEY = 'changyan_add_friend_phones';
    const SESSION_DONE_KEY = 'changyan_add_friend_session_done';
    const TAB_ID_KEY = 'changyan_add_friend_tab_id';
    const XLSX_CDN_LIST = [
        'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
        'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js',
        'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    ];
    let sheetLoaded = typeof XLSX !== 'undefined' && typeof XLSX.read === 'function';
    let sheetLoadPromise = null;
    let talkList = [];
    let talkIndex = 0;
    let phoneList = [];
    let phoneIndex = 0;
    let rateLimitRoundCount = 0;
    let successSinceCooldown = 0;
    let stableSuccessTarget = 0;
    let frequentHitCount = 0;
    let consecutiveRateLimitHits = 0;
    let running = false;
    let stopRequested = false;
    let successCount = 0;
    let failCount = 0;
    let skipCount = 0;
    let deferCount = 0;
    let deferredPhoneLog = [];
    let deferredPhoneKeys = new Set();
    let keepAliveActive = false;
    let cachedSearchInput = null;
    let visibleTextCache = { text: '', at: 0 };
    let progressBackupTimer = null;
    let lastStatsKey = '';
    let lastProgressKey = '';
    let inFlightPhone = '';
    let sessionDoneKeys = new Set();
    /** 单次 addFriendAttempt 内是否已触发过搜索回车 */
    let searchEnterUsedThisAttempt = false;

    const RATE_LIMIT_HINTS = [
        '请不要频繁点击',
        '请不要频繁',
        '搜索频繁',
        '操作频繁',
        '点击过于频繁',
        '1分钟后再试',
        '请1分钟后再试',
        '请一分钟后再试',
    ];
    const HINT_SELECTORS = [
        '.semi-toast-content',
        '.semi-toast-wrapper',
        '.semi-notification-content',
        '[role="alert"]',
        '[class*="toast"]',
        '[class*="Toast"]',
        '.semi-modal-content',
        '[class*="modal"]',
        '[class*="Modal"]',
        '[class*="notification"]',
    ];
    /** 畅言点「添加好友」后进入的申请页（非弹窗） */
    const FRIEND_APPLY_ROUTE_TITLES = ['申请添加朋友', '填写验证消息'];

    let panel = null;
    let elMiniBtn = null;
    let elTalk = null;
    let elStatus = null;
    let countdownActive = false;
    let countdownLeftSec = 0;
    let countdownTotalSec = 0;
    const COUNTDOWN_RING_LEN = 238.76;
    let elProgress = null;
    let elProgressBar = null;
    let elStats = null;
    let elBtnStart = null;
    let elBtnStop = null;
    let elBtnImport = null;
    let elBtnClear = null;
    let elFileXlsx = null;

    function isSheetReady() {
        return typeof XLSX !== 'undefined' && typeof XLSX.read === 'function';
    }

    function markSheetReady() {
        sheetLoaded = isSheetReady();
        if (sheetLoaded && elBtnImport) elBtnImport.disabled = false;
        if (sheetLoaded && elBtnClear && !running) elBtnClear.disabled = false;
        return sheetLoaded;
    }

    function loadSheetLib() {
        if (markSheetReady()) return Promise.resolve(true);
        if (sheetLoadPromise) return sheetLoadPromise;

        sheetLoadPromise = new Promise(resolve => {
            let idx = 0;
            const tryNext = () => {
                if (markSheetReady()) {
                    setStatus('Excel 库已就绪');
                    resolve(true);
                    return;
                }
                if (idx >= XLSX_CDN_LIST.length) {
                    setStatus('Excel 库加载失败，请刷新页面');
                    resolve(false);
                    return;
                }
                setStatus(`正在加载 Excel 库 (${idx + 1}/${XLSX_CDN_LIST.length})…`);
                const s = document.createElement('script');
                s.src = XLSX_CDN_LIST[idx++];
                s.onload = () => {
                    if (markSheetReady()) {
                        setStatus('Excel 库已就绪');
                        resolve(true);
                    } else {
                        tryNext();
                    }
                };
                s.onerror = tryNext;
                document.head.appendChild(s);
            };
            tryNext();
        });
        return sheetLoadPromise;
    }

    /** 后台/最小化时仍尽量准时（基于时间戳，避免被节流卡死） */
    const delay = ms => new Promise(resolve => {
        const end = Date.now() + ms;
        const tick = () => {
            const left = end - Date.now();
            if (left <= 0) return resolve();
            setTimeout(tick, Math.min(250, left));
        };
        tick();
    });

    function isOurUI(el) {
        return !!(el && el.closest && el.closest('#cy-add-friend-panel, #cy-mini-btn'));
    }

    function escapeHtml(text) {
        return String(text || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function setStatus(text) {
        if (countdownActive) return;
        if (elStatus) {
            elStatus.classList.remove('cy-waiting', 'cy-countdown-mode', 'cy-busy');
            elStatus.textContent = text;
            if (/处理:|等待|频繁|冷却|恢复|暂停/.test(text)) elStatus.classList.add('cy-busy');
        }
        updateMiniButton();
        console.log('[畅言加好友·阿陌] ' + text);
    }

    function completedCount() {
        return successCount + failCount + skipCount;
    }

    function updateProgress() {
        if (!elProgress) return;
        const total = phoneList.length;
        const completed = completedCount();
        const pct = total ? Math.min(100, Math.round((completed / total) * 100)) : 0;
        elProgress.textContent = total ? `${completed} / ${total}（${pct}%）` : '0 / 0';
        if (elProgressBar) {
            elProgressBar.style.width = (pct > 0 ? Math.max(pct, 2) : 0) + '%';
        }
    }

    function loadSessionDone() {
        try {
            const raw = sessionStorage.getItem(SESSION_DONE_KEY);
            sessionDoneKeys = new Set(JSON.parse(raw || '[]'));
        } catch (e) {
            sessionDoneKeys = new Set();
        }
    }

    function saveSessionDone() {
        try {
            sessionStorage.setItem(SESSION_DONE_KEY, JSON.stringify([...sessionDoneKeys]));
        } catch (e) {}
    }

    function markSessionDone(phone) {
        if (!phone) return;
        sessionDoneKeys.add(phoneKey(phone));
        saveSessionDone();
    }

    function isSessionDone(phone) {
        return sessionDoneKeys.has(phoneKey(phone));
    }

    function clearSessionDone() {
        sessionDoneKeys = new Set();
        try { sessionStorage.removeItem(SESSION_DONE_KEY); } catch (e) {}
    }

    async function waitUntilTabActive() {
        while (document.visibilityState === 'hidden' && running && !stopRequested) {
            saveProgress(true);
            await delay(400);
        }
        try {
            if (window.__cyAudioKeepAlive?.state === 'suspended') {
                await window.__cyAudioKeepAlive.resume();
            }
        } catch (e) {}
    }

    function clearInFlightPhone() {
        if (!inFlightPhone) return;
        inFlightPhone = '';
        saveProgress();
    }

    function setInFlightPhone(phone) {
        inFlightPhone = phone || '';
        saveProgress(true);
    }

    function phoneKey(id) {
        return /^1[3-9]\d{9}$/.test(id) ? id : String(id).toLowerCase();
    }

    /** 切页回来：申请页/提交中则续提，不重新搜索 */
    async function resumeInFlightApply(phone) {
        if (detectRateLimit()) return 'rate_limit';

        if (isFinishButtonLoading() || isSubmitInProgress()) {
            setStatus(`续等提交: ${phone}`);
            if (await waitSubmitComplete(phone)) {
                clearSearchInput();
                return 'success';
            }
            return detectRateLimit() ? 'rate_limit' : 'retry';
        }

        if (!isFriendApplyRouteOpen()) return null;

        const input = getFriendApplyRemarkInput();
        if (input && talkList.length) {
            const msg = talkList[talkIndex % talkList.length];
            await delay(PACE.friendApplySettleMs);
            if (!remarkValueMatches(readRemarkValue(input), msg)) {
                await fillFriendApplyRemark(msg);
            }
            if (!(await waitUntilRemarkFilled(msg))) return 'retry';
            await delay(PACE.remarkBeforeCompleteMs);
        }

        if (!(await clickFinishButton(PACE.friendApplyWaitMs))) return 'retry';
        if (!(await waitSubmitComplete(phone))) {
            return detectRateLimit() ? 'rate_limit' : 'retry';
        }
        clearSearchInput();
        return 'success';
    }

    async function runAttemptWithGuards(phone) {
        const startAt = Date.now();
        let attemptPromise = null;

        const launch = () => {
            attemptPromise = addFriendAttempt(phone);
            return attemptPromise;
        };
        launch();

        while (true) {
            await waitUntilTabActive();
            if (stopRequested) return 'retry';

            const raced = await Promise.race([
                attemptPromise.then(result => ({ type: 'done', result })),
                delay(800).then(() => ({ type: 'tick' })),
            ]);

            if (raced.type === 'done') return raced.result;

            if (isSubmitInProgress() || countdownActive || document.visibilityState === 'hidden') {
                continue;
            }

            if (Date.now() - startAt > PACE.stuckForceMs) {
                await forceRecoverStuck(phone);
                return 'retry';
            }
        }
    }

    function statsLine() {
        return `成功 ${successCount} · 失败 ${failCount} · 跳过 ${skipCount} · 移后 ${deferCount}`;
    }

    function updateStats(force = false) {
        const key = `${successCount}|${failCount}|${skipCount}|${deferCount}|${phoneList.length}|${phoneIndex}`;
        if (!force && key === lastStatsKey) {
            updateMiniButton();
            return;
        }
        lastStatsKey = key;
        if (elStats) {
            elStats.innerHTML = `
                <span class="cy-chip cy-chip-ok"><span class="cy-chip-icon">✓</span><span class="cy-chip-num">${successCount}</span><span class="cy-chip-label">成功</span></span>
                <span class="cy-chip cy-chip-fail"><span class="cy-chip-icon">✕</span><span class="cy-chip-num">${failCount}</span><span class="cy-chip-label">失败</span></span>
                <span class="cy-chip cy-chip-skip"><span class="cy-chip-icon">→</span><span class="cy-chip-num">${skipCount}</span><span class="cy-chip-label">跳过</span></span>
                <span class="cy-chip cy-chip-defer"><span class="cy-chip-icon">⏳</span><span class="cy-chip-num">${deferCount}</span><span class="cy-chip-label">移后</span></span>
            `;
        }
        updateProgress();
        updateMiniButton();
    }

    function updateMiniButton() {
        if (!elMiniBtn) return;
        if (countdownActive && countdownLeftSec > 0) {
            elMiniBtn.innerHTML = `<span class="cy-mini-main">${countdownLeftSec}s</span><span class="cy-mini-sub">等待</span>`;
            elMiniBtn.classList.add('cy-running', 'cy-waiting');
            return;
        }
        elMiniBtn.classList.remove('cy-waiting');
        if (running) {
            elMiniBtn.innerHTML = `<span class="cy-mini-main">畅言</span><span class="cy-mini-sub">${successCount}/${phoneList.length || 0}</span>`;
            elMiniBtn.classList.add('cy-running');
        } else {
            elMiniBtn.innerHTML = `<span class="cy-mini-main">畅言</span><span class="cy-mini-sub">阿陌</span>`;
            elMiniBtn.classList.remove('cy-running');
        }
    }

    function minimizePanel() {
        if (!panel || !elMiniBtn) return;
        panel.classList.add('cy-minimized');
        elMiniBtn.classList.add('cy-visible');
        updateMiniButton();
    }

    function restorePanel() {
        if (!panel || !elMiniBtn) return;
        panel.classList.remove('cy-minimized');
        elMiniBtn.classList.remove('cy-visible');
    }

    function startKeepAlive() {
        if (keepAliveActive) return;
        keepAliveActive = true;
        try {
            if (!window.__cyAudioKeepAlive) {
                const ctx = new (window.AudioContext || window.webkitAudioContext)();
                const osc = ctx.createOscillator();
                const gain = ctx.createGain();
                gain.gain.value = 0.001;
                osc.connect(gain);
                gain.connect(ctx.destination);
                osc.start();
                window.__cyAudioKeepAlive = ctx;
            } else {
                window.__cyAudioKeepAlive.resume();
            }
        } catch (e) {}
    }

    function stopKeepAlive() {
        keepAliveActive = false;
    }

    function invalidateVisibleTextCache() {
        visibleTextCache = { text: '', at: 0 };
    }

    function getVisibleText(force = false) {
        const now = Date.now();
        const ttl = running ? 120 : 400;
        if (!force && now - visibleTextCache.at < ttl) return visibleTextCache.text;
        const text = (document.body && document.body.innerText) || '';
        visibleTextCache = { text, at: now };
        return text;
    }

    /** 优先扫 toast/弹层，避免每次读整页 innerText */
    function collectHintText() {
        const chunks = [];
        for (const sel of HINT_SELECTORS) {
            document.querySelectorAll(sel).forEach(el => {
                if (isOurUI(el)) return;
                const t = (el.innerText || el.textContent || '').trim();
                if (t) chunks.push(t);
                const aria = (el.getAttribute('aria-label') || '').trim();
                if (aria) chunks.push(aria);
            });
        }
        return chunks.join('\n');
    }

    function textHasAny(text, hints) {
        return hints.some(h => text.includes(h));
    }

    function parseVersionParts(version) {
        return String(version || '0')
            .trim()
            .split('.')
            .map(part => parseInt(part, 10) || 0);
    }

    function compareVersion(a, b) {
        const pa = parseVersionParts(a);
        const pb = parseVersionParts(b);
        const len = Math.max(pa.length, pb.length);
        for (let i = 0; i < len; i++) {
            const da = pa[i] || 0;
            const db = pb[i] || 0;
            if (da !== db) return da - db;
        }
        return 0;
    }

    function showUpdateBlocker(requiredVersion, latestVersion) {
        const target = latestVersion || requiredVersion || SCRIPT_VERSION;
        const id = 'cy-add-friend-update-blocker';
        if (document.getElementById(id)) return;

        const blocker = document.createElement('div');
        blocker.id = id;
        blocker.style.cssText =
            'position:fixed;inset:0;z-index:2147483647;background:rgba(15,23,42,0.72);' +
            'display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box;';
        blocker.innerHTML =
            '<div style="max-width:420px;background:#fff;border-radius:16px;padding:22px 20px;' +
            'font-family:Microsoft YaHei,Segoe UI,sans-serif;color:#0f172a;box-shadow:0 24px 60px rgba(0,0,0,0.25);">' +
            '<div style="font-size:18px;font-weight:800;margin-bottom:10px;">脚本版本过低，已强制停用</div>' +
            '<div style="font-size:14px;line-height:1.7;color:#475569;margin-bottom:14px;">' +
            `当前版本 <b>${SCRIPT_VERSION}</b>，最低要求 <b>${requiredVersion}</b>，最新版本 <b>${target}</b>。<br>` +
            '请在 Tampermonkey（油猴）中打开「畅言加好友」脚本，点击「检查更新」或重新安装最新版。' +
            '</div>' +
            `<a href="${DOWNLOAD_URL}" target="_blank" rel="noopener" style="display:inline-block;padding:10px 14px;` +
            'background:#0ea5e9;color:#fff;border-radius:10px;text-decoration:none;font-weight:700;">打开更新页面</a></div>';
        document.body.appendChild(blocker);

        alert(
            `畅言加好友脚本需更新\n\n当前：${SCRIPT_VERSION}\n最低：${requiredVersion}\n最新：${target}\n\n请在 Tampermonkey 中检查更新后刷新页面。`
        );
    }

    async function checkForceUpdate() {
        try {
            const cacheBust = '?t=' + Date.now();
            const [verResp, minResp] = await Promise.all([
                fetch(VERSION_URL + cacheBust, { cache: 'no-store' }),
                fetch(MIN_VERSION_URL + cacheBust, { cache: 'no-store' }),
            ]);
            const latest = ((await verResp.text()) || '').trim();
            const minVer = ((await minResp.text()) || '').trim();
            if (minVer && compareVersion(SCRIPT_VERSION, minVer) < 0) {
                showUpdateBlocker(minVer, latest || minVer);
                return false;
            }
            if (latest && compareVersion(SCRIPT_VERSION, latest) < 0) {
                console.warn('[畅言加好友] 有新版本:', latest);
            }
        } catch (e) {
            console.warn('[畅言加好友] 版本检查失败，继续使用本地版本', e);
        }
        return true;
    }

    function lockPanelPosition() {
        if (!panel) return;
        panel.style.position = 'fixed';
        panel.style.top = '12px';
        panel.style.right = '12px';
        panel.style.left = 'auto';
        panel.style.bottom = 'auto';
        panel.style.margin = '0';
    }

    function detectRateLimit(force = false) {
        if (force) invalidateVisibleTextCache();
        const quick = collectHintText();
        if (quick && textHasAny(quick, RATE_LIMIT_HINTS)) return true;
        return textHasAny(getVisibleText(force), RATE_LIMIT_HINTS);
    }

    /** 仅在本轮已模拟回车后，才把频繁提示算作搜索触发的限频 */
    function detectSearchRateLimit(force = false) {
        if (!searchEnterUsedThisAttempt) return false;
        return detectRateLimit(force);
    }

    async function dismissStaleRateLimitToasts() {
        let dismissed = false;
        for (const sel of HINT_SELECTORS) {
            for (const el of document.querySelectorAll(sel)) {
                if (isOurUI(el)) continue;
                const t = (el.innerText || el.textContent || '').trim();
                if (!textHasAny(t, RATE_LIMIT_HINTS)) continue;
                const root = el.closest(
                    '.semi-toast, .semi-notification, [class*="toast"], [class*="Toast"], [role="alert"]'
                ) || el;
                const closeBtn = root.querySelector(
                    '.semi-toast-close, .semi-icon-close, [class*="close"], button[aria-label*="关闭"], button[aria-label*="Close"]'
                );
                if (closeBtn && !isOurUI(closeBtn)) {
                    closeBtn.click();
                    dismissed = true;
                }
            }
        }
        if (dismissed) {
            invalidateVisibleTextCache();
            await delay(350);
        }
    }

    function detectUserNotFound() {
        const quick = collectHintText();
        if (quick && quick.includes('用户不存在')) return true;
        return getVisibleText().includes('用户不存在');
    }

    function resetFrequentTracking() {
        rateLimitRoundCount = 0;
        frequentHitCount = 0;
        consecutiveRateLimitHits = 0;
    }

    /** 频繁限制：不等待，移后当前号并立即处理下一号 */
    async function onRateLimitHit(phone) {
        invalidateVisibleTextCache();
        await dismissOverlaySafely();
        await dismissStaleRateLimitToasts();
        clearSearchInput();
        clearInFlightPhone();

        let skipped = phone;
        if (phoneList.length > 1) {
            const { phone: d } = deferCurrentPhoneToEnd();
            skipped = d || phone;
            savePhoneList();
        } else {
            phoneIndex++;
        }

        const next = phoneList[phoneIndex] || '—';
        setStatus(`频繁限制，跳过 ${skipped} → 下一号 ${next}`);
        updateStats();
        saveProgress(true);
        await delay(PACE.afterSkipMs);
    }

    async function forceRecoverStuck(phone) {
        setStatus(`检测到卡住，强制恢复继续: ${phone || '—'}`);
        cachedSearchInput = null;
        invalidateVisibleTextCache();
        await dismissOverlaySafely();
        clearSearchInput();
        await delay(PACE.afterRetryMs);
    }

    function normalizeDeferredLog(log) {
        const seen = new Set();
        const out = [];
        for (const phone of log || []) {
            if (!phone) continue;
            const key = phoneKey(phone);
            if (seen.has(key)) continue;
            seen.add(key);
            out.push(phone);
        }
        return out;
    }

    function rebuildDeferredStats() {
        deferredPhoneLog = normalizeDeferredLog(deferredPhoneLog);
        deferredPhoneKeys = new Set(deferredPhoneLog.map(phoneKey));
        deferCount = deferredPhoneLog.length;
    }

    /** 号码已处理完毕（成功/失败/跳过），从移后待补扫列表移除，避免重复统计 */
    function markPhoneCompleted(phone) {
        const key = phoneKey(phone);
        if (!deferredPhoneKeys.has(key)) return;
        deferredPhoneKeys.delete(key);
        deferredPhoneLog = deferredPhoneLog.filter(p => phoneKey(p) !== key);
        deferCount = deferredPhoneLog.length;
    }

    function clearDeferredStats() {
        deferCount = 0;
        deferredPhoneLog = [];
        deferredPhoneKeys = new Set();
    }

    /** 频繁时把当前号移到队列末尾；同一号码只统计一次 */
    function recordPhoneDeferred(phone) {
        if (!phone) return false;
        const key = phoneKey(phone);
        if (deferredPhoneKeys.has(key)) return false;
        deferredPhoneKeys.add(key);
        deferredPhoneLog.push(phone);
        deferCount = deferredPhoneLog.length;
        return true;
    }

    function deferCurrentPhoneToEnd() {
        if (phoneIndex < 0 || phoneIndex >= phoneList.length) return { phone: '', isNewDefer: false };
        const phone = phoneList[phoneIndex];
        const isNewDefer = recordPhoneDeferred(phone);
        phoneList.splice(phoneIndex, 1);
        phoneList.push(phone);
        if (phoneIndex >= phoneList.length) phoneIndex = 0;
        return { phone, isNewDefer };
    }

    function parseTalkList(text) {
        return (text || '').split(/\r?\n/).map(t => t.trim()).filter(Boolean);
    }

    function saveTalks() {
        try {
            if (elTalk) localStorage.setItem(STORAGE_KEY, elTalk.value);
        } catch (e) {}
    }

    function loadTalks() {
        try {
            return localStorage.getItem(STORAGE_KEY) || '';
        } catch (e) {
            return '';
        }
    }

    function listFingerprint(list) {
        return (list || []).join('\n');
    }

    /** 去重：手机号精确匹配，畅言号忽略大小写 */
    function dedupeIdentifierList(list) {
        const seen = new Set();
        const result = [];
        for (const id of list) {
            const key = /^1[3-9]\d{9}$/.test(id) ? id : String(id).toLowerCase();
            if (seen.has(key)) continue;
            seen.add(key);
            result.push(id);
        }
        return { list: result, removed: list.length - result.length };
    }

    function getTabId() {
        try {
            let id = sessionStorage.getItem(TAB_ID_KEY);
            if (!id) {
                id = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
                sessionStorage.setItem(TAB_ID_KEY, id);
            }
            return id;
        } catch (e) {
            return 'default';
        }
    }

    function progressBackupKey() {
        return `${PROGRESS_STORAGE_KEY}_backup_${getTabId()}`;
    }

    function progressStatePayload() {
        return {
            v: 2,
            phoneIndex,
            successCount,
            failCount,
            skipCount,
            deferredPhoneLog,
            talkIndex,
            stableSuccessTarget,
            successSinceCooldown,
            frequentHitCount,
            rateLimitRoundCount,
            consecutiveRateLimitHits,
            inFlightPhone,
        };
    }

    function applyProgressState(data) {
        phoneIndex = Math.min(Math.max(0, data.phoneIndex || 0), phoneList.length);
        successCount = data.successCount || 0;
        failCount = data.failCount || 0;
        skipCount = data.skipCount || 0;
        deferredPhoneLog = normalizeDeferredLog(
            Array.isArray(data.deferredPhoneLog) ? data.deferredPhoneLog : []
        );
        rebuildDeferredStats();
        talkIndex = data.talkIndex || 0;
        stableSuccessTarget = data.stableSuccessTarget || 0;
        successSinceCooldown = data.successSinceCooldown || 0;
        frequentHitCount = data.frequentHitCount || 0;
        rateLimitRoundCount = data.rateLimitRoundCount || 0;
        consecutiveRateLimitHits = data.consecutiveRateLimitHits || 0;
        inFlightPhone = data.inFlightPhone || '';
        updateStats();
        updateStartButtonLabel();
    }

    function applyProgressData(data) {
        phoneList = data.phoneList;
        applyProgressState(data);
    }

    /** 号码列表单独存，避免每加一个好友就序列化整表 */
    function savePhoneList(flushBackup = false) {
        if (!phoneList.length) return;
        try {
            const payload = JSON.stringify(phoneList);
            sessionStorage.setItem(PHONE_LIST_STORAGE_KEY, payload);
            if (flushBackup) {
                localStorage.setItem(`${progressBackupKey()}_phones`, payload);
            }
        } catch (e) {}
    }

    function scheduleProgressBackup() {
        if (progressBackupTimer) return;
        progressBackupTimer = setTimeout(() => {
            progressBackupTimer = null;
            try {
                const state = sessionStorage.getItem(PROGRESS_STORAGE_KEY);
                const phones = sessionStorage.getItem(PHONE_LIST_STORAGE_KEY);
                if (state) localStorage.setItem(`${progressBackupKey()}_state`, state);
                if (phones) localStorage.setItem(`${progressBackupKey()}_phones`, phones);
            } catch (e) {}
        }, 400);
    }

    /** flushBackup=true：暂停/离开页面，立刻双写；否则只写 session（毫秒级） */
    function saveProgress(flushBackup = false) {
        if (!phoneList.length) return;
        try {
            const state = JSON.stringify(progressStatePayload());
            if (!flushBackup && state === lastProgressKey) return;
            lastProgressKey = state;
            sessionStorage.setItem(PROGRESS_STORAGE_KEY, state);
            if (flushBackup) {
                if (progressBackupTimer) {
                    clearTimeout(progressBackupTimer);
                    progressBackupTimer = null;
                }
                localStorage.setItem(`${progressBackupKey()}_state`, state);
                savePhoneList(true);
            } else {
                scheduleProgressBackup();
            }
        } catch (e) {}
    }

    function loadProgress() {
        try {
            let stateRaw = sessionStorage.getItem(PROGRESS_STORAGE_KEY);
            let phonesRaw = sessionStorage.getItem(PHONE_LIST_STORAGE_KEY);

            if (!phonesRaw) phonesRaw = localStorage.getItem(`${progressBackupKey()}_phones`);
            if (!stateRaw) stateRaw = localStorage.getItem(`${progressBackupKey()}_state`);

            if (!phonesRaw && stateRaw) {
                const legacy = JSON.parse(stateRaw);
                if (legacy && Array.isArray(legacy.phoneList) && legacy.phoneList.length) {
                    applyProgressData(legacy);
                    savePhoneList(true);
                    saveProgress(true);
                    return true;
                }
            }

            if (!stateRaw || !phonesRaw) {
                const raw = localStorage.getItem(PROGRESS_STORAGE_KEY);
                if (raw) {
                    sessionStorage.setItem(PROGRESS_STORAGE_KEY, raw);
                    localStorage.removeItem(PROGRESS_STORAGE_KEY);
                    return loadProgress();
                }
                return false;
            }

            const data = JSON.parse(stateRaw);
            phoneList = JSON.parse(phonesRaw);
            if (!phoneList.length) return false;

            if (data.v === 2) {
                applyProgressState(data);
            } else if (data.fingerprint === listFingerprint(phoneList)) {
                applyProgressData(data);
            } else {
                return false;
            }

            sessionStorage.setItem(PHONE_LIST_STORAGE_KEY, phonesRaw);
            return true;
        } catch (e) {
            return false;
        }
    }

    function clearProgress() {
        try {
            if (progressBackupTimer) {
                clearTimeout(progressBackupTimer);
                progressBackupTimer = null;
            }
            lastProgressKey = '';
            sessionStorage.removeItem(PROGRESS_STORAGE_KEY);
            sessionStorage.removeItem(PHONE_LIST_STORAGE_KEY);
            localStorage.removeItem(`${progressBackupKey()}_state`);
            localStorage.removeItem(`${progressBackupKey()}_phones`);
            localStorage.removeItem(PROGRESS_STORAGE_KEY);
            clearSessionDone();
            inFlightPhone = '';
        } catch (e) {}
        updateStartButtonLabel();
    }

    function bindProgressGuard() {
        if (window.__cyProgressGuard) return;
        window.__cyProgressGuard = true;
        const flush = () => {
            if (phoneList.length && (running || completedCount() > 0 || phoneIndex > 0)) {
                saveProgress(true);
            }
        };
        window.addEventListener('pagehide', flush);
        document.addEventListener('visibilitychange', () => {
            if (document.visibilityState === 'hidden') {
                flush();
                return;
            }
            if (running) {
                startKeepAlive();
                if (inFlightPhone && (isSubmitInProgress() || isFriendApplyRouteOpen())) {
                    setStatus(`已回到页面，续等: ${inFlightPhone}`);
                }
            }
        });
    }

    function updateStartButtonLabel() {
        if (!elBtnStart || running) return;
        const hasResume =
            phoneList.length > 0 &&
            phoneIndex < phoneList.length &&
            (phoneIndex > 0 || completedCount() > 0);
        elBtnStart.textContent = hasResume ? '继续' : '开始';
    }

    async function pollUntil(testFn, maxMs, interval = PACE.pollFastMs) {
        const endAt = Date.now() + maxMs;
        while (Date.now() < endAt) {
            const result = testFn();
            if (result) return result;
            await delay(interval);
        }
        return testFn();
    }

    const SEARCH_INPUT_SEL =
        'input.semi-input, input[type="text"], input[placeholder*="手机号"], input[placeholder*="畅言"]';

    function findSearchInput() {
        if (cachedSearchInput && document.contains(cachedSearchInput)) return cachedSearchInput;
        cachedSearchInput = getFriendAddSearchInput() || document.querySelector(SEARCH_INPUT_SEL);
        return cachedSearchInput;
    }

    function getFriendAddSearchInput() {
        const selectors = [
            '.wk-friendadd .wk-search-input input',
            '.wk-friendadd .wk-search-input textarea',
            '.wk-friendadd .semi-input input',
            '.wk-friendadd .semi-input textarea',
            '.wk-friendadd input.semi-input',
            '.wk-friendadd input[type="text"]',
            '.wk-friendadd input[placeholder*="手机"]',
            '.wk-friendadd input[placeholder*="号"]',
        ];
        for (const sel of selectors) {
            const el = document.querySelector(sel);
            if (el && isInputReady(el) && !isOurUI(el)) return el;
        }
        const root = document.querySelector('.wk-friendadd');
        if (!root) return null;
        const inputs = root.querySelectorAll('input[type="text"], input:not([type]), textarea');
        for (const el of inputs) {
            if (isOurUI(el) || !isInputReady(el)) continue;
            const ph = (el.getAttribute('placeholder') || '').toLowerCase();
            if (ph.includes('搜索') || ph.includes('手机') || ph.includes('号') || ph.includes('畅言')) return el;
        }
        for (const el of inputs) {
            if (!isOurUI(el) && isInputReady(el)) return el;
        }
        return null;
    }

    function isOnFriendAddPage() {
        if (document.querySelector('.wk-friendadd, .wk-search-input')) return true;
        const title = getActiveRouteTitle();
        return title.includes('添加好友');
    }

    function isFriendAddSearchInput(input) {
        if (!input) return false;
        if (input.closest?.('.wk-friendadd, .wk-search-input')) return true;
        if (isOnFriendAddPage() && input.matches('input[type="text"], input.semi-input, input:not([type])')) return true;
        return getFriendAddSearchInput() === input;
    }

    function getNativeValueSetter(input) {
        const tag = (input?.tagName || '').toLowerCase();
        const proto = tag === 'textarea' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
        return Object.getOwnPropertyDescriptor(proto, 'value')?.set;
    }

    /** v2.3 逻辑：填值 + 模拟回车搜索（只触发一次） */
    async function simulateTypingSearch(input, value) {
        if (searchEnterUsedThisAttempt) return true;
        const text = String(value || '').trim();
        if (!text || !input) return false;

        input.focus();
        const setter = getNativeValueSetter(input);
        if (setter) {
            setter.call(input, '');
            setter.call(input, text);
        } else {
            input.value = text;
        }
        input.dispatchEvent(new Event('input', { bubbles: true }));
        await delay(120);

        ['keydown', 'keypress', 'keyup'].forEach(type => {
            input.dispatchEvent(new KeyboardEvent(type, {
                key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true,
            }));
        });
        searchEnterUsedThisAttempt = true;
        setStatus(`回车搜索: ${text}`);
        await delay(500);
        return true;
    }

    /** v2.3 逻辑：填备注/话术（不回车） */
    async function simulateTyping(input, value) {
        if (!input) return;
        const text = String(value || '');
        input.focus();
        if (input.tagName.toLowerCase() === 'div') {
            document.execCommand('selectAll');
            document.execCommand('delete');
            input.textContent = text;
            input.dispatchEvent(new InputEvent('input', { bubbles: true }));
        } else {
            const setter = getNativeValueSetter(input);
            if (setter) {
                setter.call(input, '');
                setter.call(input, text);
            } else {
                input.value = text;
            }
            input.dispatchEvent(new Event('input', { bubbles: true }));
        }
        await delay(120);
    }

    async function submitSearch(input, phone) {
        const keyword = String(phone || '').trim();
        if (!keyword) return false;
        return simulateTypingSearch(input, keyword);
    }

    function normalizeUiText(el) {
        return (el?.textContent || '').replace(/\s+/g, '').trim();
    }

    function isVisibleUiEl(el) {
        if (!el || isOurUI(el)) return false;
        const r = el.getBoundingClientRect();
        return r.width > 2 && r.height > 2;
    }

    function findTextTargetInRoot(root, texts, { exact = true } = {}) {
        if (!root) return null;
        const nodes = root.querySelectorAll(
            'button, .semi-button, [role="button"], a, div, span, p, li'
        );
        let best = null;
        let bestArea = Infinity;
        for (const el of nodes) {
            if (isOurUI(el)) continue;
            const text = normalizeUiText(el);
            if (!text) continue;
            const matched = exact
                ? texts.some(t => text === String(t).replace(/\s+/g, ''))
                : texts.some(t => text.includes(String(t).replace(/\s+/g, '')));
            if (!matched) continue;
            if (!isVisibleUiEl(el)) continue;
            const r = el.getBoundingClientRect();
            const area = r.width * r.height;
            if (area < bestArea) {
                best = el;
                bestArea = area;
            }
        }
        return best;
    }

    async function humanClick(el) {
        if (!el) return false;
        try { el.scrollIntoView({ block: 'center', inline: 'center', behavior: 'instant' }); } catch (e) {
            try { el.scrollIntoView(); } catch (e2) {}
        }
        await delay(100);
        const r = el.getBoundingClientRect();
        const x = r.left + Math.max(4, r.width / 2);
        const y = r.top + Math.max(4, r.height / 2);
        const opts = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y };
        for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
            const Cls = type.startsWith('pointer') ? PointerEvent : MouseEvent;
            el.dispatchEvent(new Cls(type, opts));
        }
        let node = el;
        for (let i = 0; i < 16 && node; i++) {
            const props = getReactProps(node);
            if (props?.onClick) {
                try { props.onClick({}); } catch (e) {}
                try { props.onClick({ nativeEvent: {}, target: el, currentTarget: node }); } catch (e) {}
            }
            node = node.parentElement;
        }
        try { el.click(); } catch (e) {}
        return true;
    }

    function findAddFriendButton() {
        const footerSelectors = [
            '.wk-userinfo-footer-sendbutton',
            '.wk-userInfo-footer-sendbutton',
            '.wk-userinfo-footer',
            '.wk-userInfo-footer',
            '.wk-userinfo',
            '.wk-userInfo',
            '.wk-userInfo-content',
            '.wk-userinfo-content',
        ];
        for (const sel of footerSelectors) {
            const roots = document.querySelectorAll(sel);
            for (const root of roots) {
                const hit = findTextTargetInRoot(root, ['添加好友'], { exact: true });
                if (hit) return hit;
            }
        }
        return findTextTargetInRoot(document.body, ['添加好友'], { exact: true });
    }

    async function waitInputStable(getInput, stableMs, maxMs) {
        const endAt = Date.now() + (maxMs || PACE.friendApplyWaitMs);
        let last = null;
        let stableSince = 0;
        while (Date.now() < endAt) {
            const input = getInput();
            if (input && isInputReady(input)) {
                if (input === last) {
                    if (Date.now() - stableSince >= (stableMs || 600)) return input;
                } else {
                    last = input;
                    stableSince = Date.now();
                }
            } else {
                last = null;
                stableSince = 0;
            }
            await delay(PACE.pollFastMs);
        }
        const input = getInput();
        return input && isInputReady(input) ? input : null;
    }

    async function waitFriendApplyDialogReady(phone) {
        setStatus(`等待申请页加载: ${phone}`);
        const endAt = Date.now() + PACE.friendApplyWaitMs;
        while (Date.now() < endAt) {
            if (isFriendApplyRouteOpen()) {
                const input = await waitInputStable(getFriendApplyRemarkInput, 700, PACE.friendApplyWaitMs);
                if (input) {
                    await delay(PACE.friendApplySettleMs);
                    return input;
                }
            }
            if (detectRateLimit()) return 'rate_limit';
            await delay(PACE.pollFastMs);
        }
        return null;
    }
    async function clickAddFriendButton(phone) {
        const endAt = Date.now() + PACE.searchProfileWaitMs;
        while (Date.now() < endAt) {
            if (isFriendApplyRouteOpen()) {
                await delay(PACE.friendApplySettleMs);
                return true;
            }

            const target = findAddFriendButton();
            if (target) {
                setStatus(`等待资料页: ${phone}`);
                await delay(PACE.profilePanelSettleMs);
                setStatus(`点击添加好友: ${phone}`);
                await humanClick(target);
                await delay(PACE.afterAddFriendClickMs);

                const ready = await waitFriendApplyDialogReady(phone);
                if (ready === 'rate_limit') return 'rate_limit';
                if (ready) return true;

                const retryTarget = findAddFriendButton();
                if (retryTarget) {
                    setStatus(`再次点击添加好友: ${phone}`);
                    await delay(PACE.actionSettleMs);
                    await humanClick(retryTarget);
                    await delay(PACE.afterAddFriendClickMs);
                    const retryReady = await waitFriendApplyDialogReady(phone);
                    if (retryReady === 'rate_limit') return 'rate_limit';
                    if (retryReady) return true;
                }
                return false;
            }

            if (findExistingFriendPanel()) return false;
            const blocked = checkRateLimitOrNotFound();
            if (blocked) return blocked;
            setStatus(`等待资料页「添加好友」: ${phone}`);
            await delay(PACE.pollNormalMs);
        }
        return null;
    }

    function clearSearchInput(input) {
        const el = input || findSearchInput();
        if (!el) return;
        setInputValue(el, '');
    }

    function isInputReady(input) {
        if (!input || !document.contains(input)) return false;
        if (input.disabled || input.readOnly) return false;
        const r = input.getBoundingClientRect();
        return r.width > 4 && r.height > 4;
    }

    function readInputValue(input) {
        return input.tagName.toLowerCase() === 'div'
            ? input.textContent.trim()
            : input.value.trim();
    }

    async function setInputValueAndEnter(input, value) {
        input.focus();
        await delay(PACE.typingMs);
        setInputValue(input, value);
        await delay(PACE.typingMs);
        ['keydown', 'keypress', 'keyup'].forEach(type => {
            input.dispatchEvent(new KeyboardEvent(type, {
                key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
            }));
        });
        await delay(PACE.typingAfterMs);
    }

    function getReactFiber(el) {
        if (!el) return null;
        for (const key of Object.keys(el)) {
            if (key.startsWith('__reactFiber$') || key.startsWith('__reactInternalInstance$')) {
                return el[key];
            }
        }
        return null;
    }

    function getReactProps(el) {
        const fiber = getReactFiber(el);
        return fiber?.memoizedProps || fiber?.pendingProps || null;
    }

    function syncReactInputValue(input, value) {
        let el = input;
        for (let i = 0; i < 12 && el; i++) {
            const props = getReactProps(el);
            if (props?.onChange) {
                try { props.onChange(value); } catch (e) { /* ignore */ }
                try { props.onChange(value, { target: input, currentTarget: input }); } catch (e) { /* ignore */ }
            }
            el = el.parentElement;
        }
    }

    function setInputValue(input, value) {
        input.focus();
        if (input.tagName.toLowerCase() === 'div') {
            document.execCommand('selectAll');
            document.execCommand('delete');
            input.textContent = value;
            input.dispatchEvent(new InputEvent('input', { bubbles: true }));
            syncReactInputValue(input, value);
            return;
        }
        const tag = input.tagName.toLowerCase();
        const proto = tag === 'textarea' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
        const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
        if (setter) setter.call(input, value);
        else input.value = value;
        input.dispatchEvent(new InputEvent('input', {
            bubbles: true,
            cancelable: true,
            data: value,
            inputType: 'insertFromPaste',
        }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        syncReactInputValue(input, value);
    }

    /** 回车后轮询到用户资料页：优先等「添加好友」，勿与已是好友面板混淆 */
    async function waitSearchOutcome(phone, waitFriendPanel) {
        const endAt = Date.now() + PACE.searchProfileWaitMs;
        while (Date.now() < endAt) {
            const blocked = checkRateLimitOrNotFound({ searchPhase: true });
            if (blocked) return blocked;
            if (findAddFriendButton()) {
                setStatus(`等待资料页加载: ${phone}`);
                await delay(PACE.profilePanelSettleMs);
                return null;
            }
            if (findExistingFriendPanel()) {
                return await handleAlreadyFriend(phone, false);
            }
            await delay(PACE.pollFastMs);
        }
            if (findAddFriendButton()) {
                setStatus(`等待资料页加载: ${phone}`);
                await delay(PACE.profilePanelSettleMs);
                return null;
            }
        if (findExistingFriendPanel()) {
            return await handleAlreadyFriend(phone, waitFriendPanel);
        }
        return checkRateLimitOrNotFound({ searchPhase: true });
    }

    function waitForSelector(selector, maxMs = 40 * PACE.pollFastMs, interval = PACE.pollFastMs) {
        return pollUntil(() => document.querySelector(selector), maxMs, interval);
    }

    function waitForButton(keyword, role = 'button', maxMs = 60 * PACE.pollNormalMs, interval = PACE.pollNormalMs, root = document) {
        return pollUntil(() => {
            return Array.from(root.querySelectorAll(`button, div[role="${role}"]`))
                .find(el => !isOurUI(el) && (el.innerText || '').includes(keyword)) || null;
        }, maxMs, interval);
    }

    function findExistingFriendPanel() {
            if (findAddFriendButton()) {
                setStatus(`等待资料页加载: ${phone}`);
                await delay(PACE.profilePanelSettleMs);
                return null;
            }

        const sendBtns = Array.from(document.querySelectorAll('button, div[role="button"], a'))
            .filter(el => {
                if (isOurUI(el)) return false;
                const t = (el.innerText || '').trim();
                return t === '发送消息' || t.includes('发送消息');
            });

        for (const btn of sendBtns) {
            let node = btn;
            for (let i = 0; i < 12 && node; i++) {
                const text = node.innerText || '';
                if (
                    text.includes('发送消息') &&
                    (text.includes('设置备注') || text.includes('解除好友关系') || text.includes('拉入黑名单'))
                ) {
                    const r = node.getBoundingClientRect();
                    if (r.width > 80 && r.height > 120) return node;
                }
                node = node.parentElement;
            }
        }
        return null;
    }

    async function waitPanelClosed(maxMs = PACE.panelCloseMs) {
        await pollUntil(() => !findExistingFriendPanel(), maxMs);
        return !findExistingFriendPanel();
    }

    async function closeExistingFriendPanel(panel) {
        panel = panel || findExistingFriendPanel();
        if (!panel) return true;

        const root = panel.closest('.semi-modal, .modal, [class*="modal"]') || panel;
        const box = root.getBoundingClientRect();

        const scopedClose = root.querySelector(
            '.semi-modal-close, .semi-icons-close, [class*="modal-close"], [class*="Modal-close"]'
        );
        if (scopedClose) {
            scopedClose.click();
            await delay(PACE.panelCloseMinMs);
            return waitPanelClosed();
        }

        const candidates = Array.from(root.querySelectorAll('button, span, div, i, svg')).filter(el => {
            if (isOurUI(el)) return false;
            const r = el.getBoundingClientRect();
            if (r.width < 8 || r.height < 8 || r.width > 48 || r.height > 48) return false;
            if (r.top > box.top + 72 || r.left > box.left + 72) return false;
            const label = (el.innerText || '').trim();
            const cls = String(el.className || '');
            const aria = el.getAttribute('aria-label') || '';
            return (
                label === '×' || label === '✕' || label === 'X' ||
                /close/i.test(cls) || aria.includes('关闭') || aria.includes('Close')
            );
        });

        if (candidates.length) {
            candidates.sort((a, b) => {
                const ra = a.getBoundingClientRect();
                const rb = b.getBoundingClientRect();
                return (ra.top + ra.left) - (rb.top + rb.left);
            });
            (candidates[0].closest('button') || candidates[0]).click();
            await delay(PACE.panelCloseMinMs);
            return waitPanelClosed();
        }

        return false;
    }

    async function dismissOverlaySafely() {
        const friendPanel = findExistingFriendPanel();
        if (friendPanel) return closeExistingFriendPanel(friendPanel);
        return clickConfirmComplete();
    }

    /** 已是好友：发送消息+设置备注等面板 → 点 X 关闭 → 跳过当前号 */
    async function handleAlreadyFriend(phone, waitPanel) {
        let friendPanel = findExistingFriendPanel();
        if (!friendPanel && waitPanel) {
            friendPanel = await pollUntil(() => findExistingFriendPanel(), 12 * PACE.pollFastMs);
        }
        if (!friendPanel) return null;

        await closeExistingFriendPanel(friendPanel);
        if (findExistingFriendPanel()) {
            await delay(PACE.panelCloseRetryMs);
            await closeExistingFriendPanel();
        }
        if (findExistingFriendPanel()) {
            const back = document.querySelector(
                '.wk-chat-conversation-header-back, [class*="header-back"], [class*="nav-back"]'
            );
            if (back && !isOurUI(back)) {
                back.click();
                await waitPanelClosed();
            }
        }
        setStatus('已是好友，已跳过: ' + phone);
        return 'already_friend';
    }

    /** 频率限制 / 用户不存在（搜索阶段限频须已回车） */
    function checkRateLimitOrNotFound({ searchPhase = false } = {}) {
        if (searchPhase ? detectSearchRateLimit() : detectRateLimit()) return 'rate_limit';
        if (detectUserNotFound()) return 'not_found';
        return null;
    }

    async function checkAfterAction(phone, waitFriendPanel) {
        const blocked = checkRateLimitOrNotFound();
        if (blocked) return blocked;
        return await handleAlreadyFriend(phone, waitFriendPanel);
    }

    function getActiveRouteTitle() {
        const titles = document.querySelectorAll('.wk-viewqueueheader-content-title');
        for (let i = titles.length - 1; i >= 0; i--) {
            const t = (titles[i].textContent || '').trim();
            if (t) return t;
        }
        return '';
    }

    function isFriendApplyRouteOpen() {
        if (document.querySelector('.wk-friendapply')) return true;
        const title = getActiveRouteTitle();
        return FRIEND_APPLY_ROUTE_TITLES.some(t => title.includes(t));
    }

    function getFriendApplyRemarkInput() {
        const selectors = [
            '.wk-friendapply-content-message textarea',
            '.wk-friendapply-content-message input[type="text"]',
            '.wk-friendapply-content-message .semi-input textarea',
            '.wk-friendapply-content-message .semi-input input',
            '.wk-friendapply-content-message .semi-input-textarea textarea',
            '.wk-friendapply-content-message div[contenteditable="true"]',
            '.wk-friendapply textarea.semi-textarea',
            '.wk-friendapply textarea',
            '.wk-friendapply input[type="text"]',
            '.wk-friendapply .semi-input-textarea textarea',
            '.wk-friendapply .semi-input textarea',
            '.wk-friendapply div[contenteditable="true"]',
        ];
        for (const sel of selectors) {
            const el = document.querySelector(sel);
            if (el && isInputReady(el)) return el;
        }
        const scope = document.querySelector('.wk-friendapply, .wk-friendapply-content-message');
        if (scope) {
            for (const el of scope.querySelectorAll('textarea, input[type="text"], div[contenteditable="true"]')) {
                if (isOurUI(el) || !isInputReady(el)) continue;
                const ph = (el.getAttribute('placeholder') || '').toLowerCase();
                if (ph.includes('验证') || ph.includes('留言') || ph.includes('申请') || ph.includes('备注') || ph.includes('朋友')) {
                    return el;
                }
            }
            for (const el of scope.querySelectorAll('textarea, input[type="text"], div[contenteditable="true"]')) {
                if (!isOurUI(el) && isInputReady(el)) return el;
            }
        }
        return null;
    }

    async function waitFriendApplyRemarkInput(maxMs) {
        const endAt = Date.now() + (maxMs || PACE.remarkFillWaitMs);
        while (Date.now() < endAt) {
            const input = getFriendApplyRemarkInput();
            if (input && isInputReady(input)) return input;
            await delay(PACE.pollFastMs);
        }
        return getFriendApplyRemarkInput();
    }

    function readRemarkValue(input) {
        if (!input) return '';
        if (input.tagName.toLowerCase() === 'div') return (input.textContent || input.innerText || '').trim();
        return (input.value || '').trim();
    }

    function remarkValueMatches(current, expected) {
        const cur = (current || '').trim();
        const exp = (expected || '').trim();
        return !!exp && (cur === exp || cur.includes(exp));
    }

    function syncFriendApplyRemarkState(remark) {
        const input = getFriendApplyRemarkInput();
        const scopes = [
            document.querySelector('.wk-friendapply-content-message'),
            document.querySelector('.wk-friendapply'),
            input,
        ].filter(Boolean);

        const seen = new Set();
        const tryProps = (el, depth = 0) => {
            if (!el || depth > 20 || seen.has(el)) return;
            seen.add(el);
            const props = getReactProps(el);
            if (props?.onChange) {
                try { props.onChange(remark); } catch (e) { /* ignore */ }
                try { props.onChange(remark, { target: { value: remark }, currentTarget: { value: remark } }); } catch (e) { /* ignore */ }
            }
            if (props?.onInput) {
                try { props.onInput(remark); } catch (e) { /* ignore */ }
                try { props.onInput({ target: { value: remark }, currentTarget: { value: remark } }); } catch (e) { /* ignore */ }
            }
            if (props?.onValueChange) {
                try { props.onValueChange(remark); } catch (e) { /* ignore */ }
            }
            if (props?.onMessage) {
                try { props.onMessage(remark); } catch (e) { /* ignore */ }
            }
            if (el.parentElement) tryProps(el.parentElement, depth + 1);
            for (const child of el.children || []) tryProps(child, depth + 1);
        };
        scopes.forEach(scope => tryProps(scope));
    }

    async function typeIntoApplyRemark(input, remark) {
        if (!input || !remark) return false;
        const text = String(remark).trim();
        input.focus();
        await delay(PACE.typingMs);

        const tryFill = async () => {
            setInputValue(input, text);
            syncReactInputValue(input, text);
            syncFriendApplyRemarkState(text);
            await delay(PACE.remarkSettleMs);
            const fresh = getFriendApplyRemarkInput() || input;
            return remarkValueMatches(readRemarkValue(fresh), text);
        };

        if (await tryFill()) return true;

        input.focus();
        try {
            document.execCommand('selectAll', false, null);
            document.execCommand('insertText', false, text);
        } catch (e) { /* ignore */ }
        input.dispatchEvent(new InputEvent('input', {
            bubbles: true, cancelable: true, data: text, inputType: 'insertText',
        }));
        syncReactInputValue(input, text);
        syncFriendApplyRemarkState(text);
        await delay(PACE.remarkSettleMs);
        if (remarkValueMatches(readRemarkValue(getFriendApplyRemarkInput() || input), text)) return true;

        setInputValue(input, '');
        await delay(60);
        for (const ch of text) {
            const next = readRemarkValue(input) + ch;
            setInputValue(input, next);
            input.dispatchEvent(new InputEvent('input', {
                bubbles: true, cancelable: true, data: ch, inputType: 'insertText',
            }));
            syncReactInputValue(input, next);
            await delay(70);
        }
        syncFriendApplyRemarkState(text);
        await delay(PACE.remarkCheckMs);
        return remarkValueMatches(readRemarkValue(getFriendApplyRemarkInput() || input), text);
    }

    async function waitUntilRemarkFilled(expected, maxMs) {
        const text = String(expected || '').trim();
        if (!text) return false;
        const endAt = Date.now() + (maxMs || PACE.remarkFillWaitMs);
        while (Date.now() < endAt) {
            const input = getFriendApplyRemarkInput() || findRemarkInputSync();
            if (input && remarkValueMatches(readRemarkValue(input), text)) return true;
            await delay(PACE.pollFastMs);
        }
        const input = getFriendApplyRemarkInput() || findRemarkInputSync();
        return !!(input && remarkValueMatches(readRemarkValue(input), text));
    }

    async function fillFriendApplyRemark(remark) {
        const text = String(remark || '').trim();
        if (!text) return false;
        const attempts = PACE.remarkFillRetries || 3;
        for (let i = 1; i <= attempts; i++) {
            const input = await waitFriendApplyRemarkInput(PACE.remarkFillWaitMs);
            if (!input) {
                await delay(PACE.remarkSettleMs);
                continue;
            }
            setStatus(i > 1 ? `重填验证消息 (${i}/${attempts})…` : `填写验证消息…`);
            await delay(PACE.friendApplySettleMs);
            await simulateTyping(input, text);
            await delay(PACE.remarkSettleMs);
            if (remarkValueMatches(readRemarkValue(getFriendApplyRemarkInput() || input), text)) {
                return true;
            }
            if (await typeIntoApplyRemark(input, text)) return true;
            await delay(PACE.remarkSettleMs);
            if (await waitUntilRemarkFilled(text, 2500)) return true;
        }
        return waitUntilRemarkFilled(text, PACE.remarkFillWaitMs);
    }

    function findFinishButton() {
        if (document.querySelector('.wk-friendapply')) {
            const headers = document.querySelectorAll('.wk-viewqueueheader');
            for (let i = headers.length - 1; i >= 0; i--) {
                const btn = headers[i].querySelector(
                    '.wk-viewqueueheader-content-action button, .wk-viewqueueheader-content-action .semi-button'
                );
                if (btn && !isOurUI(btn)) return btn;
            }
        }
        const headers = document.querySelectorAll('.wk-viewqueueheader');
        for (const header of headers) {
            const title = header.querySelector('.wk-viewqueueheader-content-title')?.textContent?.trim() || '';
            if (!FRIEND_APPLY_ROUTE_TITLES.some(t => title.includes(t))) continue;
            const btn = header.querySelector(
                '.wk-viewqueueheader-content-action button, .wk-viewqueueheader-content-action .semi-button'
            );
            if (!btn || isOurUI(btn)) continue;
            const text = (btn.textContent || '').replace(/\s+/g, '').trim();
            if (text === '完成') return btn;
        }
        return Array.from(document.querySelectorAll('button, .semi-button, [role="button"]'))
            .find(el => !isOurUI(el) && (el.textContent || '').replace(/\s+/g, '').trim() === '完成') || null;
    }

    function isFinishButtonEnabled(btn) {
        if (!btn) return false;
        if (btn.disabled || btn.getAttribute('aria-disabled') === 'true') return false;
        if (btn.classList.contains('semi-button-disabled')) return false;
        if (btn.classList.contains('semi-button-loading')) return false;
        return true;
    }

    function nodeHasSpinner(el) {
        if (!el) return false;
        if (el.classList.contains('semi-button-loading')) return true;
        if (el.getAttribute('aria-busy') === 'true') return true;
        if (el.querySelector(
            '.semi-spin, .semi-spin-wrapper, .semi-icon-loading, .semi-icon-spin, ' +
            '[class*="spin"], [class*="Spin"], [class*="loading"], [class*="Loading"]'
        )) return true;
        return false;
    }

    function isFinishButtonLoading() {
        const btn = findFinishButton();
        if (nodeHasSpinner(btn)) return true;
        const action = document.querySelector('.wk-viewqueueheader-content-action');
        if (nodeHasSpinner(action)) return true;
        if (isFriendApplyRouteOpen() && btn && !isFinishButtonEnabled(btn) && nodeHasSpinner(btn.closest('.wk-viewqueueheader-content-action'))) {
            return true;
        }
        return false;
    }

    function isSubmitInProgress() {
        return isFinishButtonLoading() || isFriendApplyRouteOpen() || isAddFriendModalOpen();
    }

    async function clickFinishButton(maxWait) {
        const endAt = Date.now() + (maxWait || PACE.confirmWaitMs + 2000);
        while (Date.now() < endAt) {
            const btn = findFinishButton();
            if (btn && isFinishButtonEnabled(btn)) {
                await delay(PACE.remarkBeforeCompleteMs);
                const props = getReactProps(btn);
                try {
                    if (props?.onClick) props.onClick({});
                    else btn.click();
                } catch (e) {
                    btn.click();
                }
                await delay(PACE.confirmClickMs);
                for (let i = 0; i < 10; i++) {
                    if (isFinishButtonLoading() || !isFriendApplyRouteOpen()) return true;
                    await delay(PACE.pollFastMs);
                }
                return true;
            }
            await delay(PACE.pollFastMs);
        }
        return false;
    }

    /** 等待「完成」提交结束：loading 消失且离开申请页后再继续 */
    async function waitSubmitComplete(phone) {
        if (!isSubmitInProgress()) return true;

        await delay(PACE.submitSettleMinMs);
        const startAt = Date.now();
        const endAt = startAt + PACE.submitWaitMs;
        let sawLoading = false;

        while (Date.now() < endAt) {
            if (detectRateLimit()) return false;

            const loading = isFinishButtonLoading();
            const onApply = isFriendApplyRouteOpen();
            const modalOpen = isAddFriendModalOpen();
            if (loading) sawLoading = true;

            if (phone) {
                setStatus(loading ? `提交中: ${phone}（请稍候）` : `等待提交: ${phone}`);
            }

            if (!loading && !onApply && !modalOpen) {
                const elapsed = Date.now() - startAt;
                if (sawLoading || elapsed >= PACE.submitSettleMinMs + 600) {
                    await delay(450);
                    if (!isSubmitInProgress()) return true;
                }
            }

            await delay(PACE.pollNormalMs);
        }

        return !isSubmitInProgress();
    }

    function findRemarkInputSync() {
        const friendApplyInput = getFriendApplyRemarkInput();
        if (friendApplyInput && isInputReady(friendApplyInput)) return friendApplyInput;

        const container = document.querySelector('.semi-modal-content, .modal-content, .popup-content') || document.body;
        const textarea = container.querySelector('textarea.semi-textarea, textarea');
        if (textarea) return textarea;

        const inputs = container.querySelectorAll('input[type="text"]');
        for (const inp of inputs) {
            const ph = (inp.getAttribute('placeholder') || '').toLowerCase();
            if (ph.includes('备注') || ph.includes('留言') || ph.includes('验证')) return inp;
        }

        return container.querySelector('div[contenteditable="true"]') || null;
    }

    function isAddFriendModalOpen() {
        const modal = document.querySelector('.semi-modal-content, .modal-content, .popup-content');
        if (!modal) return false;
        return Array.from(modal.querySelectorAll('button, div[role="button"]'))
            .some(el => !isOurUI(el) && ['完成', '确定', '确认'].some(k => (el.innerText || '').includes(k)));
    }

    async function waitForRemarkInput(phone) {
        await delay(PACE.friendApplySettleMs);
        const endAt = Date.now() + Math.max(PACE.actionWaitMs, PACE.friendApplyWaitMs);
        while (Date.now() < endAt) {
            const blocked = checkRateLimitOrNotFound();
            if (blocked) return blocked;
            if (findExistingFriendPanel()) {
                return await handleAlreadyFriend(phone, false);
            }
            if (isFriendApplyRouteOpen()) {
                const applyInput = await waitInputStable(getFriendApplyRemarkInput, 700, PACE.friendApplyWaitMs);
                if (applyInput) {
                    await delay(PACE.friendApplySettleMs);
                    return applyInput;
                }
            }
            const input = findRemarkInputSync();
            if (input && isInputReady(input)) {
                await delay(PACE.friendApplySettleMs);
                return input;
            }
            await delay(PACE.pollFastMs);
        }
        return null;
    }

    async function fillRemarkInput(input, msg) {
        if (isFriendApplyRouteOpen() || input?.closest?.('.wk-friendapply')) {
            return fillFriendApplyRemark(msg);
        }
        await typeIntoApplyRemark(input, msg);
        return waitUntilRemarkFilled(msg, PACE.remarkFillWaitMs);
    }

    async function checkRemarkFilled(input, expected) {
        const fresh = getFriendApplyRemarkInput() || input;
        if (remarkValueMatches(readInputValue(fresh), expected)) return true;
        await delay(PACE.remarkCheckMs);
        const again = getFriendApplyRemarkInput() || input;
        return remarkValueMatches(readInputValue(again), expected);
    }

    async function clickConfirmComplete() {
        try {
            await delay(PACE.remarkBeforeCompleteMs);
            const modal = document.querySelector('.semi-modal-content, .modal-content, .popup-content');
            const root = modal || document.body;
            const keywords = ['完成', '确定', '确认'];
            const endAt = Date.now() + 20 * PACE.pollNormalMs;
            while (Date.now() < endAt) {
                for (const kw of keywords) {
                    const btn = Array.from(root.querySelectorAll('button, div[role="button"]'))
                        .find(el => !isOurUI(el) && (el.innerText || '').includes(kw) && !el.disabled);
                    if (btn) {
                        btn.click();
                        await delay(PACE.confirmClickMs);
                        return true;
                    }
                }
                await delay(PACE.pollNormalMs);
            }
            return false;
        } catch (e) {
            return false;
        }
    }

    /**
     * success      → 添加成功，下一号
     * not_found    → 用户不存在，下一号
     * already_friend → 已是好友（点X关闭），下一号
     * rate_limit   → 频繁限制，当前号移到最后，立即下一号
     * retry        → 未完成添加，清弹窗后重试当前号
     */
    async function addFriendAttempt(phone) {
        await waitUntilTabActive();

        if (isSessionDone(phone)) {
            clearInFlightPhone();
            return 'success';
        }

        if (inFlightPhone && inFlightPhone !== phone && (isSubmitInProgress() || isFriendApplyRouteOpen())) {
            setStatus(`等待 ${inFlightPhone} 提交完成…`);
            await waitSubmitComplete(inFlightPhone);
        }

        if (inFlightPhone === phone) {
            const resumed = await resumeInFlightApply(phone);
            if (resumed) {
                clearInFlightPhone();
                if (resumed === 'success') markSessionDone(phone);
                return resumed;
            }
        }

        if (isFinishButtonLoading()) {
            setStatus(`上一号仍在提交，等待…`);
            if (!(await waitSubmitComplete(phone))) {
                if (detectRateLimit()) return 'rate_limit';
                return 'retry';
            }
        }

        setInFlightPhone(phone);
        searchEnterUsedThisAttempt = false;

        const input = findSearchInput() || await waitForSelector(SEARCH_INPUT_SEL);
        if (!input) return 'retry';
        cachedSearchInput = input;

        await dismissStaleRateLimitToasts();

        if (!(await submitSearch(input, phone))) return 'retry';
        if (detectSearchRateLimit(true)) return 'rate_limit';
        invalidateVisibleTextCache();
        let outcome = await waitSearchOutcome(phone, true);
        if (outcome) return outcome;

        const clickResult = await clickAddFriendButton(phone);
        if (clickResult === 'rate_limit' || clickResult === 'not_found') return clickResult;
        if (!clickResult) {
            if (findExistingFriendPanel()) {
                return await handleAlreadyFriend(phone, false);
            }
            setStatus(`未找到「添加好友」按钮，重试: ${phone}`);
            return 'retry';
        }

        const remarkResult = await waitForRemarkInput(phone);
        if (typeof remarkResult === 'string') return remarkResult;
        if (!remarkResult) {
            outcome = await checkAfterAction(phone, true);
            if (outcome) return outcome;
            setStatus(`未进入验证消息页，重试: ${phone}`);
            return 'retry';
        }
        if (talkList.length === 0) return 'retry';

        const msg = talkList[talkIndex++ % talkList.length];
        setStatus(`填写验证消息: ${msg.length > 18 ? msg.slice(0, 18) + '…' : msg}`);
        let filled = false;
        if (isFriendApplyRouteOpen()) {
            await delay(PACE.friendApplySettleMs);
            filled = await fillFriendApplyRemark(msg);
        } else {
            await delay(PACE.friendApplySettleMs);
            filled = await fillRemarkInput(remarkResult, msg);
        }
        if (!filled) {
            setStatus(`验证消息未填入，重试: ${phone}`);
            return 'retry';
        }

        if (!(await waitUntilRemarkFilled(msg))) {
            setStatus(`话术未确认填入，重试: ${phone}`);
            return 'retry';
        }
        await delay(PACE.remarkBeforeCompleteMs);

        let confirmed = false;
        if (isFriendApplyRouteOpen()) {
            confirmed = await clickFinishButton(PACE.friendApplyWaitMs);
        } else {
            confirmed = await clickConfirmComplete();
        }
        if (!confirmed) {
            setStatus(`未找到「完成」按钮，重试: ${phone}`);
            return 'retry';
        }

        if (!(await waitSubmitComplete(phone))) {
            if (detectRateLimit()) return 'rate_limit';
            setStatus(`提交未完成，重试: ${phone}`);
            return 'retry';
        }

        outcome = checkRateLimitOrNotFound();
        if (outcome) return outcome;

        await delay(PACE.typingAfterMs);
        clearSearchInput(input);
        clearInFlightPhone();
        setStatus('添加成功: ' + phone);
        return 'success';
    }

    async function startLoop() {
        if (running) return;

        talkList = parseTalkList(elTalk ? elTalk.value : '');
        saveTalks();

        if (talkList.length === 0) {
            alert('请先在下方编辑框输入话术，每行一条');
            return;
        }
        if (phoneList.length === 0) {
            alert('请先导入号码 Excel');
            return;
        }
        if (phoneIndex >= phoneList.length) {
            phoneIndex = 0;
            successCount = 0;
            failCount = 0;
            skipCount = 0;
            talkIndex = 0;
            resetFrequentTracking();
            clearDeferredStats();
            clearSessionDone();
            clearProgress();
        }

        running = true;
        stopRequested = false;
        if (phoneIndex === 0 && completedCount() === 0) {
            resetFrequentTracking();
        }
        updateStats();
        startKeepAlive();
        elBtnStart.disabled = true;
        elBtnStop.disabled = false;
        elBtnImport.disabled = true;
        if (elBtnClear) elBtnClear.disabled = true;
        if (elTalk) elTalk.disabled = true;

        if (phoneIndex > 0 || completedCount() > 0) {
            const nextPhone = phoneList[phoneIndex] || '';
            setStatus(`已完成 ${completedCount()}/${phoneList.length}，继续: ${nextPhone}${deferCount ? `（移后 ${deferCount} 个待补扫）` : ''}`);
        } else {
            setStatus('开始加好友（后台可继续）...');
        }

        while (phoneIndex < phoneList.length) {
            if (stopRequested) {
                saveProgress(true);
                setStatus(`已暂停 | 已完成 ${completedCount()}/${phoneList.length} | 下次继续: ${phoneList[phoneIndex] || '—'} | ${statsLine()}`);
                break;
            }

            const phone = phoneList[phoneIndex];
            updateStats();

            await waitUntilTabActive();

            if (isSessionDone(phone)) {
                phoneIndex++;
                saveProgress();
                continue;
            }

            await dismissStaleRateLimitToasts();

            if (isFinishButtonLoading()) {
                setStatus(`等待上一号提交完成…`);
                await waitSubmitComplete('');
            }

            setStatus(`处理: ${phone}`);

            const result = await runAttemptWithGuards(phone);

            if (result === 'success') {
                markPhoneCompleted(phone);
                markSessionDone(phone);
                clearInFlightPhone();
                successCount++;
                phoneIndex++;
                rateLimitRoundCount = 0;
                frequentHitCount = 0;
                consecutiveRateLimitHits = 0;
                updateStats();
                setStatus(`${statsLine()} | 已完成: ${phone}`);
                await delay(PACE.afterSuccessMs);
            } else if (result === 'not_found') {
                markPhoneCompleted(phone);
                markSessionDone(phone);
                clearInFlightPhone();
                failCount++;
                phoneIndex++;
                rateLimitRoundCount = 0;
                frequentHitCount = 0;
                consecutiveRateLimitHits = 0;
                updateStats();
                setStatus(`用户不存在，跳过: ${phone} | ${statsLine()}`);
                await delay(PACE.afterNotFoundMs);
            } else if (result === 'already_friend') {
                markPhoneCompleted(phone);
                markSessionDone(phone);
                clearInFlightPhone();
                skipCount++;
                phoneIndex++;
                rateLimitRoundCount = 0;
                frequentHitCount = 0;
                consecutiveRateLimitHits = 0;
                updateStats();
                setStatus(`已是好友，跳过: ${phone} | ${statsLine()}`);
                await delay(PACE.afterSkipMs);
            } else if (result === 'rate_limit') {
                await onRateLimitHit(phone);
            } else {
                const failHint = searchEnterUsedThisAttempt ? '搜索后未完成' : '填号后未完成';
                setStatus(`${failHint}，重试当前号: ${phone}`);
                await dismissOverlaySafely();
                const inp = findSearchInput();
                const cur = inp ? readInputValue(inp) : '';
                if (cur !== phone) clearSearchInput(inp);
                await delay(PACE.afterRetryMs);
            }

            saveProgress(stopRequested);

            if (stopRequested) {
                saveProgress(true);
                setStatus(`已暂停 | 已完成 ${completedCount()}/${phoneList.length} | 下次继续: ${phoneList[phoneIndex] || '—'} | ${statsLine()}`);
                break;
            }
        }

        if (!stopRequested && phoneIndex >= phoneList.length) {
            setStatus(`全部完成 | ${statsLine()}`);
            clearInFlightPhone();
            clearSessionDone();
            clearProgress();
        }

        running = false;
        stopKeepAlive();
        elBtnStart.disabled = false;
        elBtnStop.disabled = true;
        elBtnImport.disabled = false;
        if (elBtnClear) elBtnClear.disabled = false;
        if (elTalk) elTalk.disabled = false;
        updateStartButtonLabel();
        updateMiniButton();
    }

    function stopLoop() {
        stopRequested = true;
        saveProgress(true);
        setStatus(
            phoneList.length
                ? `已暂停并保存 | 已完成 ${completedCount()}/${phoneList.length} | 下次继续: ${phoneList[phoneIndex] || '—'}`
                : '已暂停'
        );
    }

    function clearAllData() {
        if (running) return;
        if (phoneList.length && !confirm('确定清空已导入的号码和进度？')) return;
        phoneList = [];
        phoneIndex = 0;
        successCount = 0;
        failCount = 0;
        skipCount = 0;
        talkIndex = 0;
        resetFrequentTracking();
        clearDeferredStats();
        clearProgress();
        updateStats();
        updateStartButtonLabel();
        setStatus('已清空导入列表与进度');
    }

    /** 从单元格文本/数字中识别中国大陆手机号 */
    function normalizePhone(raw, allowEmbedded) {
        if (raw === null || raw === undefined) return null;
        let s = String(raw).trim();
        if (!s) return null;

        if (/^\d+\.?\d*[eE][+\-]?\d+$/.test(s)) {
            const num = Number(s);
            if (Number.isFinite(num)) s = String(Math.round(num));
        }

        if (/^1[3-9]\d{9}$/.test(s)) return s;

        let digits = s.replace(/\D/g, '');
        if (digits.startsWith('86') && digits.length === 13) digits = digits.slice(2);
        if (/^1[3-9]\d{9}$/.test(digits)) return digits;

        if (allowEmbedded) {
            const embedded = s.match(/(?:\+?86[-\s]?)?(1[3-9]\d{9})/);
            if (embedded) return embedded[1];
        }

        return null;
    }

    /** 系统内部 UID（32位十六进制等），不是畅言号 */
    function isSystemUid(s) {
        if (!s) return false;
        const t = String(s).trim();
        if (/^[0-9a-fA-F]{32}$/.test(t)) return true;
        if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/i.test(t)) return true;
        return false;
    }

    /** 姓名、地址、纯英文单词等，不应作为号码导入 */
    function isSkippableCellContent(raw) {
        const s = String(raw ?? '').trim();
        if (!s) return true;
        if (/[\u4e00-\u9fff]/.test(s)) return true;
        if (/[,，、；;|｜]/.test(s) && s.length > 12) return true;
        if (/^(男|女|m|f|male|female|yes|no|true|false)$/i.test(s)) return true;
        if (/^\d{1,2}$/.test(s)) return true;
        if (/^[a-zA-Z]{2,}$/.test(s) && !/\d/.test(s)) return true;
        if (/^(http|https|www\.)/i.test(s)) return true;
        if (/^\d{4}[-/年]\d{1,2}([-/月]\d{1,2})?/.test(s)) return true;
        return false;
    }

    /**
     * 严格识别畅言号（整格匹配）：须为纯数字或字母+数字，排除纯英文、UID
     */
    function normalizeChangyanIdStrict(raw) {
        if (raw === null || raw === undefined) return null;
        let s = String(raw).trim();
        if (!s || isSkippableCellContent(s)) return null;

        if (normalizePhone(s, false)) return null;
        if (isSystemUid(s)) return null;

        if (/^\d+\.?\d*[eE][+\-]?\d+$/.test(s)) {
            const num = Number(s);
            if (Number.isFinite(num)) s = String(Math.round(num));
        }

        if (isSystemUid(s)) return null;

        if (/^[a-zA-Z0-9]+$/.test(s) && /[a-zA-Z]/.test(s) && /\d/.test(s) && s.length >= 3 && s.length <= 32) {
            return s;
        }

        const labeled = s.match(/(?:畅言号|畅言账号)[:：\s]*([a-zA-Z0-9]{3,32})/);
        if (labeled && !isSystemUid(labeled[1]) && (/\d/.test(labeled[1]) || /^\d+$/.test(labeled[1]))) {
            return labeled[1];
        }

        if (/^\d+$/.test(s) && s.length >= 3 && s.length <= 20) {
            return s;
        }

        return null;
    }

    function normalizeIdentifierStrict(raw) {
        if (isSkippableCellContent(raw)) return null;
        return normalizePhone(raw, false) || normalizeChangyanIdStrict(raw);
    }

    function normalizeFromCell(raw, colKind) {
        if (isSkippableCellContent(raw)) return null;
        if (colKind === 'phone') return normalizePhone(raw, true);
        if (colKind === 'changyan') return normalizeChangyanIdStrict(raw);
        return normalizeIdentifierStrict(raw);
    }

    const COL_SKIP_RE = /姓名|名字|联系人|客户|地址|公司|部门|微信|备注|序号|编号|地区|省份|城市|邮箱|性别|年龄|生日|创建|更新|uid|uuid|id$/i;
    const COL_PHONE_RE = /手机|电话|号码|mobile|phone|联系电话|手机号|tel/i;
    const COL_CYID_RE = /畅言号|畅言账号|畅言|账号|account|用户名|user/i;

    function classifyColumnHeader(cellText) {
        const t = String(cellText || '').trim();
        if (!t) return 'unknown';
        if (/^uid$/i.test(t) || COL_SKIP_RE.test(t)) return 'skip';
        if (COL_PHONE_RE.test(t)) return 'phone';
        if (COL_CYID_RE.test(t)) return 'changyan';
        return 'unknown';
    }

    function sheetToRows(sheet) {
        if (!sheet || !sheet['!ref']) return [];
        const range = XLSX.utils.decode_range(sheet['!ref']);
        const rows = [];
        for (let r = range.s.r; r <= range.e.r; r++) {
            const rowCells = [];
            for (let c = range.s.c; c <= range.e.c; c++) {
                rowCells.push(getCellDisplay(sheet, r, c));
            }
            if (rowCells.some(v => String(v).trim())) rows.push(rowCells);
        }
        return rows;
    }

    function findHeaderRowIndex(rows) {
        let best = -1;
        let bestScore = 0;
        const limit = Math.min(8, rows.length);
        for (let i = 0; i < limit; i++) {
            let score = 0;
            for (const cell of rows[i]) {
                const kind = classifyColumnHeader(cell);
                if (kind === 'phone' || kind === 'changyan' || kind === 'skip') score += 2;
                else if (kind === 'unknown' && String(cell).trim()) score += 0.5;
            }
            const hasDataId = rows[i].some(c => normalizeIdentifierStrict(c));
            if (hasDataId) score -= 3;
            if (score > bestScore) {
                bestScore = score;
                best = i;
            }
        }
        return bestScore >= 2 ? best : -1;
    }

    function buildColumnPlan(headerCells) {
        const targetCols = [];
        const colKinds = {};
        headerCells.forEach((cell, idx) => {
            const kind = classifyColumnHeader(cell);
            if (kind === 'phone' || kind === 'changyan') {
                targetCols.push(idx);
                colKinds[idx] = kind;
            }
        });
        return { targetCols, colKinds };
    }

    function findBestDataColumn(rows, skipRow) {
        const scores = {};
        rows.forEach((row, ri) => {
            if (ri === skipRow) return;
            row.forEach((cell, ci) => {
                if (normalizeIdentifierStrict(cell)) {
                    scores[ci] = (scores[ci] || 0) + 1;
                }
            });
        });
        let bestCol = -1;
        let bestScore = 0;
        Object.entries(scores).forEach(([c, n]) => {
            if (n > bestScore) {
                bestScore = n;
                bestCol = +c;
            }
        });
        return bestScore >= 1 ? bestCol : -1;
    }

    function extractIdFromRow(rowCells, plan) {
        if (plan.mode === 'single') {
            const cell = rowCells[plan.singleCol];
            return normalizeIdentifierStrict(cell);
        }

        if (plan.targetCols.length) {
            for (const ci of plan.targetCols) {
                const id = normalizeFromCell(rowCells[ci], plan.colKinds[ci] || 'unknown');
                if (id) return id;
            }
            return null;
        }

        for (const cell of rowCells) {
            const id = normalizeIdentifierStrict(cell);
            if (id) return id;
        }
        return null;
    }

    function getCellDisplay(sheet, r, c) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[addr];
        if (!cell) return '';
        if (cell.w != null && String(cell.w).trim()) return cell.w;
        if (cell.v !== undefined && cell.v !== null) return cell.v;
        return '';
    }

    /**
     * 智能导入：识别表头列（手机/畅言号），跳过姓名地址；
     * 无表头时按单列或每行取一条，避免误扫英文和杂项。
     */
    function extractPhonesFromSheet(sheet) {
        const rows = sheetToRows(sheet);
        if (!rows.length) return [];

        const headerIdx = findHeaderRowIndex(rows);
        let plan;

        if (headerIdx >= 0) {
            const built = buildColumnPlan(rows[headerIdx]);
            plan = {
                headerIdx,
                targetCols: built.targetCols,
                colKinds: built.colKinds,
                mode: 'multi',
                singleCol: -1,
            };
            if (!plan.targetCols.length) {
                const bestCol = findBestDataColumn(rows, headerIdx);
                if (bestCol >= 0) {
                    plan.mode = 'single';
                    plan.singleCol = bestCol;
                }
            }
        } else {
            const bestCol = findBestDataColumn(rows, -1);
            if (bestCol >= 0) {
                plan = { headerIdx: -1, targetCols: [], colKinds: {}, mode: 'single', singleCol: bestCol };
            } else {
                plan = { headerIdx: -1, targetCols: [], colKinds: {}, mode: 'multi', singleCol: -1 };
            }
        }

        const list = [];
        rows.forEach((row, ri) => {
            if (ri === plan.headerIdx) return;
            const id = extractIdFromRow(row, plan);
            if (id) list.push(id);
        });

        return list;
    }

    async function importPhonesFromFile(file) {
        if (!file) return;

        setStatus('正在准备 Excel 库…');
        if (elBtnImport) elBtnImport.disabled = true;
        const ok = await loadSheetLib();
        if (elBtnImport && !running) elBtnImport.disabled = false;
        if (!ok) {
            alert('Excel 库加载失败，请检查网络后刷新页面重试');
            return;
        }

        const reader = new FileReader();
        reader.onload = evt => {
            try {
                const data = new Uint8Array(evt.target.result);
                const wb = XLSX.read(data, { type: 'array', cellDates: false });
                const sheet = wb.Sheets[wb.SheetNames[0]];
                const rawList = extractPhonesFromSheet(sheet);
                const { list: phones, removed } = dedupeIdentifierList(rawList);
                if (phones.length === 0) {
                    alert('未找到有效手机号或畅言号\n已自动跳过姓名、地址、纯英文等列\n支持：11位手机号、纯数字畅言号、字母+数字畅言号');
                    return;
                }
                phoneList = phones;
                phoneIndex = 0;
                successCount = 0;
                failCount = 0;
                skipCount = 0;
                talkIndex = 0;
                resetFrequentTracking();
                clearDeferredStats();
                clearProgress();
                savePhoneList(true);
                saveProgress(true);
                updateStats();
                updateStartButtonLabel();
                let msg = '已导入 ' + phones.length + ' 个号码（手机号/畅言号）';
                if (removed > 0) msg += '，已去除 ' + removed + ' 个重复';
                setStatus(msg);
            } catch (e) {
                alert('读取 Excel 失败: ' + e);
            }
        };
        reader.onerror = () => alert('读取文件失败，请重新选择');
        reader.readAsArrayBuffer(file);
    }

    function createPanel() {
        if (document.getElementById('cy-add-friend-panel')) return;

        const style = document.createElement('style');
        style.textContent = `
            @keyframes cy-shimmer {
                0% { background-position: 200% center; }
                100% { background-position: -200% center; }
            }
            @keyframes cy-pulse-soft {
                0%, 100% { opacity: 0.55; transform: scale(1); }
                50% { opacity: 1; transform: scale(1.04); }
            }
            @keyframes cy-float {
                0%, 100% { transform: translateY(0); }
                50% { transform: translateY(-2px); }
            }
            #cy-add-friend-panel {
                position: fixed; top: 12px; right: 12px; z-index: 99999;
                width: 388px; background: rgba(255,255,255,0.97);
                border-radius: 22px; overflow: hidden;
                border: 1px solid rgba(125,211,252,0.65);
                backdrop-filter: blur(14px);
                box-shadow:
                    0 28px 64px rgba(14,165,233,0.16),
                    0 10px 28px rgba(15,23,42,0.08),
                    inset 0 1px 0 rgba(255,255,255,0.95);
                font-family: "Segoe UI", "Microsoft YaHei UI", "PingFang SC", system-ui, sans-serif;
                font-size: 13px; color: #0f172a;
            }
            #cy-add-friend-panel.cy-minimized { display: none !important; }
            #cy-add-friend-panel .cy-head {
                position: relative; display: flex; align-items: center; justify-content: space-between;
                padding: 17px 17px 16px; overflow: hidden; user-select: none;
                background: linear-gradient(135deg, #e0f2fe 0%, #bae6fd 24%, #7dd3fc 58%, #38bdf8 100%);
                border-bottom: 1px solid rgba(255,255,255,0.5);
            }
            #cy-add-friend-panel .cy-head::before {
                content: ""; position: absolute; inset: 0; pointer-events: none;
                background:
                    radial-gradient(circle at 90% 12%, rgba(255,255,255,0.62) 0%, transparent 44%),
                    radial-gradient(circle at 8% 88%, rgba(255,255,255,0.32) 0%, transparent 38%),
                    radial-gradient(circle at 52% 120%, rgba(14,165,233,0.12) 0%, transparent 55%);
            }
            #cy-add-friend-panel .cy-head::after {
                content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 1px;
                background: linear-gradient(90deg, transparent, rgba(255,255,255,0.85), transparent);
            }
            #cy-add-friend-panel .cy-head-main { position: relative; z-index: 1; min-width: 0; }
            #cy-add-friend-panel .cy-head-title {
                display: flex; align-items: center; gap: 8px;
                font-weight: 800; font-size: 19px; line-height: 1.15; letter-spacing: 0.02em;
                color: #075985; text-shadow: 0 1px 0 rgba(255,255,255,0.5);
            }
            #cy-add-friend-panel .cy-head-icon {
                display: inline-flex; align-items: center; justify-content: center;
                width: 28px; height: 28px; border-radius: 10px;
                font-size: 14px; font-weight: 900; color: #0284c7;
                background: rgba(255,255,255,0.72);
                border: 1px solid rgba(255,255,255,0.9);
                box-shadow: 0 2px 8px rgba(14,165,233,0.14);
                animation: cy-float 3s ease-in-out infinite;
            }
            #cy-add-friend-panel .cy-head-sub {
                display: flex; align-items: center; flex-wrap: wrap; gap: 6px;
                margin-top: 7px; font-size: 11px; color: #0369a1;
            }
            #cy-add-friend-panel .cy-badge {
                display: inline-flex; align-items: center; padding: 2px 9px;
                border-radius: 999px; font-size: 10px; font-weight: 700;
                color: #0369a1; background: rgba(255,255,255,0.78);
                border: 1px solid rgba(255,255,255,0.92);
                box-shadow: 0 1px 3px rgba(14,165,233,0.14);
            }
            #cy-add-friend-panel .cy-ver {
                padding: 2px 7px; border-radius: 999px; font-size: 10px; font-weight: 600;
                color: #0c4a6e; background: rgba(255,255,255,0.45);
                border: 1px solid rgba(255,255,255,0.6);
            }
            #cy-add-friend-panel .cy-head-btn {
                position: relative; z-index: 1;
                background: rgba(255,255,255,0.58); border: 1px solid rgba(255,255,255,0.82);
                color: #0369a1; width: 36px; height: 36px; border-radius: 12px; cursor: pointer;
                line-height: 1; font-size: 20px; font-weight: 500; flex-shrink: 0;
                box-shadow: 0 2px 10px rgba(14,165,233,0.14);
                transition: background .15s ease, transform .12s ease, box-shadow .15s ease;
            }
            #cy-add-friend-panel .cy-head-btn:hover {
                background: rgba(255,255,255,0.88); transform: translateY(-1px);
                box-shadow: 0 4px 14px rgba(14,165,233,0.18);
            }
            #cy-add-friend-panel .cy-body {
                padding: 15px; display: flex; flex-direction: column; gap: 12px;
                background: linear-gradient(180deg, #eff8ff 0%, #f8fafc 42%, #ffffff 100%);
            }
            #cy-add-friend-panel .cy-card {
                background: rgba(255,255,255,0.94); border: 1px solid #dbeafe;
                border-radius: 18px; padding: 15px 15px 14px;
                box-shadow: 0 3px 14px rgba(14,165,233,0.07);
            }
            #cy-add-friend-panel .cy-progress-head {
                display: flex; align-items: baseline; justify-content: space-between;
                gap: 8px; margin-bottom: 12px;
            }
            #cy-add-friend-panel .cy-progress-label {
                font-size: 11px; font-weight: 700; color: #64748b;
                letter-spacing: 0.08em; text-transform: uppercase;
            }
            #cy-add-friend-panel .cy-progress {
                font-size: 18px; color: #0c4a6e; font-weight: 800; font-variant-numeric: tabular-nums;
            }
            #cy-add-friend-panel .cy-stats {
                display: grid; grid-template-columns: repeat(4, 1fr); gap: 7px; margin-bottom: 12px;
            }
            #cy-add-friend-panel .cy-chip {
                position: relative; display: flex; flex-direction: column; align-items: center;
                justify-content: center; gap: 2px; padding: 9px 3px 8px; border-radius: 14px;
                border: 1px solid transparent; min-height: 58px; min-width: 0;
                transition: transform .14s ease, box-shadow .14s ease;
            }
            #cy-add-friend-panel .cy-chip:hover { transform: translateY(-1px); }
            #cy-add-friend-panel .cy-chip-icon {
                font-size: 10px; font-weight: 800; line-height: 1; opacity: 0.75;
            }
            #cy-add-friend-panel .cy-chip-num {
                font-size: 18px; font-weight: 800; line-height: 1; font-variant-numeric: tabular-nums;
            }
            #cy-add-friend-panel .cy-chip-label {
                font-size: 10px; font-weight: 600; opacity: 0.88;
            }
            #cy-add-friend-panel .cy-chip-ok {
                background: linear-gradient(180deg,#ecfdf5,#d1fae5); border-color: #a7f3d0; color: #047857;
                box-shadow: 0 2px 8px rgba(16,185,129,0.08);
            }
            #cy-add-friend-panel .cy-chip-fail {
                background: linear-gradient(180deg,#fef2f2,#fee2e2); border-color: #fecaca; color: #b91c1c;
                box-shadow: 0 2px 8px rgba(239,68,68,0.07);
            }
            #cy-add-friend-panel .cy-chip-skip {
                background: linear-gradient(180deg,#f8fafc,#f1f5f9); border-color: #e2e8f0; color: #475569;
                box-shadow: 0 2px 8px rgba(100,116,139,0.06);
            }
            #cy-add-friend-panel .cy-chip-defer {
                background: linear-gradient(180deg,#fffbeb,#fef3c7); border-color: #fde68a; color: #b45309;
                box-shadow: 0 2px 8px rgba(245,158,11,0.08);
            }
            #cy-add-friend-panel .cy-progress-track {
                height: 9px; background: #e0f2fe; border-radius: 99px; overflow: hidden;
                box-shadow: inset 0 1px 3px rgba(14,165,233,0.1);
            }
            #cy-add-friend-panel .cy-progress-bar {
                height: 100%; width: 0%; border-radius: 99px;
                background: linear-gradient(90deg, #7dd3fc, #0ea5e9, #0284c7, #0ea5e9);
                background-size: 200% 100%;
                animation: cy-shimmer 2.8s linear infinite;
                box-shadow: 0 0 12px rgba(14,165,233,0.38);
                transition: width .4s cubic-bezier(.4,0,.2,1);
            }
            #cy-add-friend-panel .cy-field-card {
                background: rgba(255,255,255,0.96); border: 1px solid #dbeafe;
                border-radius: 18px; padding: 13px 14px;
                box-shadow: 0 2px 10px rgba(14,165,233,0.06);
                box-sizing: border-box; min-width: 0;
            }
            #cy-add-friend-panel .cy-label {
                display: block; margin-bottom: 8px; color: #0369a1;
                font-size: 11px; font-weight: 700; letter-spacing: 0.05em;
            }
            #cy-add-friend-panel textarea {
                width: 100%; box-sizing: border-box; min-height: 84px; resize: vertical;
                border: 1px solid #bfdbfe; border-radius: 14px; padding: 11px 13px;
                font-size: 12px; line-height: 1.58; font-family: inherit;
                background: linear-gradient(180deg,#f8fbff,#f0f9ff);
                color: #0f172a;
                transition: border-color .15s ease, box-shadow .15s ease, background .15s ease;
            }
            #cy-add-friend-panel textarea:focus {
                outline: none; border-color: #38bdf8; background: #fff;
                box-shadow: 0 0 0 3px rgba(56,189,248,0.22);
            }
            #cy-add-friend-panel .cy-btns-wrap { display: flex; flex-direction: column; gap: 9px; }
            #cy-add-friend-panel .cy-btns-row { display: flex; gap: 9px; }
            #cy-add-friend-panel button.cy-btn {
                flex: 1; min-width: 0; padding: 12px 10px; border: none;
                border-radius: 14px; cursor: pointer; color: #fff; font-size: 13px;
                font-weight: 700; font-family: inherit; letter-spacing: 0.02em;
                box-shadow: 0 4px 12px rgba(15,23,42,0.1);
                transition: transform .12s ease, filter .12s ease, box-shadow .12s ease;
            }
            #cy-add-friend-panel button.cy-btn:hover:not(:disabled) {
                filter: brightness(1.04); transform: translateY(-2px);
                box-shadow: 0 6px 16px rgba(15,23,42,0.13);
            }
            #cy-add-friend-panel button.cy-btn:active:not(:disabled) { transform: translateY(0); }
            #cy-add-friend-panel button.cy-btn:disabled:not(.cy-btn-stop) {
                opacity: 0.46; cursor: not-allowed; transform: none; box-shadow: none;
            }
            #cy-add-friend-panel .cy-btn-import {
                background: linear-gradient(180deg, #93c5fd, #0ea5e9);
                box-shadow: 0 4px 14px rgba(14,165,233,0.3);
            }
            #cy-add-friend-panel .cy-btn-clear { background: linear-gradient(180deg,#94a3b8,#64748b); }
            #cy-add-friend-panel .cy-btn-start {
                background: linear-gradient(180deg,#4ade80,#16a34a);
                box-shadow: 0 4px 14px rgba(22,163,74,0.24);
            }
            #cy-add-friend-panel button.cy-btn.cy-btn-stop {
                background: linear-gradient(180deg,#fb7185,#e11d48) !important;
                box-shadow: 0 4px 14px rgba(225,29,72,0.22) !important;
            }
            #cy-add-friend-panel button.cy-btn.cy-btn-stop:disabled {
                opacity: 0.48 !important; cursor: not-allowed !important;
                transform: none !important; box-shadow: none !important;
            }
            #cy-add-friend-panel .cy-status {
                padding: 14px 16px; background: rgba(255,255,255,0.96);
                border: 1px solid #dbeafe; border-radius: 18px;
                font-size: 13px; color: #334155;
                min-height: 64px; word-break: break-all; line-height: 1.62;
                font-weight: 500; box-shadow: 0 2px 10px rgba(14,165,233,0.06);
            }
            #cy-add-friend-panel .cy-status.cy-busy {
                border-color: #7dd3fc; background: linear-gradient(180deg,#f0f9ff,#e0f2fe);
                color: #0369a1; font-weight: 600;
            }
            #cy-add-friend-panel .cy-status.cy-countdown-mode {
                text-align: center; border-color: #fbbf24;
                background: linear-gradient(180deg, #fffbeb 0%, #fef3c7 55%, #fde68a 100%);
                color: #92400e; min-height: 128px;
                display: flex; flex-direction: column; align-items: center; justify-content: center;
                box-shadow: 0 6px 20px rgba(245,158,11,0.14);
            }
            #cy-add-friend-panel .cy-status-count-label {
                font-size: 12px; font-weight: 700; color: #b45309; margin-bottom: 8px;
                letter-spacing: 0.03em;
            }
            #cy-add-friend-panel .cy-countdown-visual {
                position: relative; width: 96px; height: 96px;
                display: flex; align-items: center; justify-content: center;
            }
            #cy-add-friend-panel .cy-countdown-ring {
                position: absolute; inset: 0; width: 100%; height: 100%;
                transform: rotate(-90deg);
            }
            #cy-add-friend-panel .cy-ring-bg {
                fill: none; stroke: rgba(251,191,36,0.28); stroke-width: 6;
            }
            #cy-add-friend-panel .cy-ring-fg {
                fill: none; stroke: #f59e0b; stroke-width: 6; stroke-linecap: round;
                transition: stroke-dashoffset .85s linear;
                filter: drop-shadow(0 0 4px rgba(245,158,11,0.35));
            }
            #cy-add-friend-panel .cy-status-count-sec {
                position: relative; z-index: 1;
                font-size: 32px; font-weight: 800; line-height: 1;
                color: #d97706; font-variant-numeric: tabular-nums;
            }
            #cy-add-friend-panel .cy-status-count-unit {
                font-size: 13px; font-weight: 700; margin-left: 2px; color: #b45309;
            }
            #cy-add-friend-panel .cy-status-count-sub {
                margin-top: 10px; font-size: 12px; color: #92400e; opacity: 0.92;
                max-width: 100%; overflow: hidden; text-overflow: ellipsis;
                padding: 0 8px;
            }
            #cy-mini-btn {
                position: fixed; left: 12px; bottom: 72px; z-index: 9999;
                width: 54px; height: 54px; border-radius: 50%; border: none;
                background: linear-gradient(145deg, #93c5fd, #0ea5e9);
                color: #fff; cursor: pointer;
                box-shadow: 0 10px 26px rgba(14,165,233,0.34);
                display: none; flex-direction: column; align-items: center; justify-content: center;
                user-select: none; padding: 0; line-height: 1.05;
                transition: transform .15s ease, box-shadow .15s ease;
            }
            #cy-mini-btn:hover { transform: scale(1.05); }
            #cy-mini-btn.cy-visible { display: flex; }
            #cy-mini-btn.cy-running {
                box-shadow: 0 0 0 4px rgba(125,211,252,0.38), 0 10px 26px rgba(14,165,233,0.34);
            }
            #cy-mini-btn.cy-waiting {
                background: linear-gradient(145deg, #fcd34d, #f59e0b);
                box-shadow: 0 0 0 5px rgba(251,191,36,0.28), 0 10px 26px rgba(245,158,11,0.3);
                animation: cy-pulse-soft 1.6s ease-in-out infinite;
            }
            #cy-mini-btn .cy-mini-main { font-size: 13px; font-weight: 800; }
            #cy-mini-btn .cy-mini-sub { font-size: 8px; margin-top: 2px; opacity: 0.96; }
        `;
        document.head.appendChild(style);

        panel = document.createElement('div');
        panel.id = 'cy-add-friend-panel';

        const head = document.createElement('div');
        head.className = 'cy-head';
        head.innerHTML = `
            <div class="cy-head-main">
                <div class="cy-head-title"><span class="cy-head-icon">✦</span>畅言加好友</div>
                <div class="cy-head-sub">
                    <span class="cy-badge">阿陌专用</span>
                    <span class="cy-ver">v${SCRIPT_VERSION}</span>
                    <span>后台稳定版</span>
                </div>
            </div>
            <button type="button" class="cy-head-btn" id="cy-panel-minimize" title="最小化">−</button>
        `;

        const body = document.createElement('div');
        body.className = 'cy-body';

        const progressCard = document.createElement('div');
        progressCard.className = 'cy-card';

        const progressHead = document.createElement('div');
        progressHead.className = 'cy-progress-head';
        const progressLabel = document.createElement('span');
        progressLabel.className = 'cy-progress-label';
        progressLabel.textContent = '已完成';
        elProgress = document.createElement('div');
        elProgress.className = 'cy-progress';
        elProgress.textContent = '0 / 0';
        progressHead.append(progressLabel, elProgress);

        const progressTrack = document.createElement('div');
        progressTrack.className = 'cy-progress-track';
        elProgressBar = document.createElement('div');
        elProgressBar.className = 'cy-progress-bar';
        progressTrack.appendChild(elProgressBar);

        elStats = document.createElement('div');
        elStats.className = 'cy-stats';
        updateStats();

        progressCard.append(progressHead, elStats, progressTrack);

        const talkCol = document.createElement('div');
        talkCol.className = 'cy-field-card';
        const talkLabel = document.createElement('label');
        talkLabel.className = 'cy-label';
        talkLabel.textContent = '话术（每行一条）';
        elTalk = document.createElement('textarea');
        elTalk.placeholder = '你好，很高兴认识你\n方便加个好友吗';
        elTalk.value = loadTalks();
        elTalk.addEventListener('blur', saveTalks);
        talkCol.append(talkLabel, elTalk);

        const btnsWrap = document.createElement('div');
        btnsWrap.className = 'cy-btns-wrap';

        const btnRow1 = document.createElement('div');
        btnRow1.className = 'cy-btns-row';
        elBtnImport = Object.assign(document.createElement('button'), {
            type: 'button', className: 'cy-btn cy-btn-import', textContent: '导入 Excel'
        });
        elBtnClear = Object.assign(document.createElement('button'), {
            type: 'button', className: 'cy-btn cy-btn-clear', textContent: '清空'
        });
        btnRow1.append(elBtnImport, elBtnClear);

        const btnRow2 = document.createElement('div');
        btnRow2.className = 'cy-btns-row';
        elBtnStart = Object.assign(document.createElement('button'), {
            type: 'button', className: 'cy-btn cy-btn-start', textContent: '开始'
        });
        elBtnStop = Object.assign(document.createElement('button'), {
            type: 'button', className: 'cy-btn cy-btn-stop', textContent: '暂停', disabled: true
        });
        btnRow2.append(elBtnStart, elBtnStop);
        btnsWrap.append(btnRow1, btnRow2);

        elStatus = document.createElement('div');
        elStatus.className = 'cy-status';
        elStatus.textContent = '请先导入号码并填写话术';

        body.append(progressCard, talkCol, btnsWrap, elStatus);
        panel.append(head, body);
        document.body.appendChild(panel);

        elMiniBtn = document.createElement('button');
        elMiniBtn.id = 'cy-mini-btn';
        elMiniBtn.type = 'button';
        elMiniBtn.title = '畅言加好友 · 阿陌专用';
        elMiniBtn.innerHTML = '<span class="cy-mini-main">畅言</span><span class="cy-mini-sub">阿陌</span>';
        elMiniBtn.onclick = () => restorePanel();
        document.body.appendChild(elMiniBtn);

        elFileXlsx = Object.assign(document.createElement('input'), {
            type: 'file', accept: '.xlsx,.xls', style: 'display:none'
        });
        document.body.appendChild(elFileXlsx);

        elBtnImport.onclick = () => elFileXlsx.click();
        elFileXlsx.onchange = e => {
            importPhonesFromFile(e.target.files[0]);
            e.target.value = '';
        };
        elBtnStart.onclick = () => startLoop();
        elBtnStop.onclick = stopLoop;
        elBtnClear.onclick = () => clearAllData();

        document.getElementById('cy-panel-minimize').onclick = e => {
            e.stopPropagation();
            minimizePanel();
        };

        lockPanelPosition();
        window.addEventListener('resize', lockPanelPosition);

        setStatus('面板已就绪，可最小化后台运行');
    }

    async function boot() {
        getTabId();
        loadSessionDone();
        bindProgressGuard();
        const ok = await checkForceUpdate();
        if (!ok) return;
        createPanel();
        lockPanelPosition();
        if (loadProgress()) {
            setStatus(`已恢复进度，已完成 ${completedCount()}/${phoneList.length}，继续: ${phoneList[phoneIndex] || '—'}（点「继续」开始）`);
        }
        loadSheetLib().then(sheetOk => {
            if (sheetOk && !phoneList.length) setStatus('Excel 库已就绪，可导入文件');
            else if (!sheetOk) setStatus('Excel 库加载失败，请刷新页面');
        });
    }

    function startPanelWatchdog() {
        if (window.__cyPanelWatchdog) return;
        window.__cyPanelWatchdog = true;
        const ensurePanel = () => {
            if (!document.getElementById('cy-add-friend-panel')) createPanel();
            lockPanelPosition();
            if (running && !keepAliveActive) startKeepAlive();
        };
        document.addEventListener('visibilitychange', () => {
            if (document.visibilityState === 'visible') ensurePanel();
        });
        setInterval(ensurePanel, 8000);
    }

    if (document.readyState !== 'loading') boot();
    else window.addEventListener('DOMContentLoaded', boot);

    startPanelWatchdog();

})();
