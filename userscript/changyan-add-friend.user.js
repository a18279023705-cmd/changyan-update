// ==UserScript==
// @name         畅言加好友 阿陌专用 后台稳定版
// @namespace    http://tampermonkey.net/
// @version      9.7
// @description  畅言加好友阿陌专用，完善重试/跳过/已是好友判定逻辑
// @match        *://web.rvtqh.com/*
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @grant        none
// @run-at       document-end
// @updateURL    https://raw.githubusercontent.com/a18279023705-cmd/changyan-update/main/userscript/changyan-add-friend.meta.js
// @downloadURL  https://raw.githubusercontent.com/a18279023705-cmd/changyan-update/main/userscript/changyan-add-friend.user.js
// ==/UserScript==

(function () {
    'use strict';

    const STORAGE_KEY = 'changyan_add_friend_talks';
    const PROGRESS_STORAGE_KEY = 'changyan_add_friend_progress';
    const DELAY_STORAGE_KEY = 'changyan_add_friend_delay';
    const DEFAULT_DELAY_MIN = 1;
    const DEFAULT_DELAY_MAX = 5;
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
    let running = false;
    let stopRequested = false;
    let successCount = 0;
    let failCount = 0;
    let skipCount = 0;
    let keepAliveTimer = null;

    let panel = null;
    let elMiniBtn = null;
    let elTalk = null;
    let elStatus = null;
    let elProgress = null;
    let elProgressBar = null;
    let elStats = null;
    let elBtnStart = null;
    let elBtnStop = null;
    let elBtnImport = null;
    let elBtnClear = null;
    let elDelayMin = null;
    let elDelayMax = null;
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

    function setStatus(text) {
        if (elStatus) elStatus.textContent = text;
        updateMiniButton();
        console.log('[畅言加好友·阿陌] ' + text);
    }

    function updateProgress() {
        if (!elProgress) return;
        const total = phoneList.length;
        const current = total ? Math.min(phoneIndex + 1, total) : 0;
        const pct = total ? Math.min(100, Math.round((phoneIndex / total) * 100)) : 0;
        elProgress.textContent = total ? `${current} / ${total}（${pct}%）` : '0 / 0';
        if (elProgressBar) {
            elProgressBar.style.width = (pct > 0 ? Math.max(pct, 2) : 0) + '%';
        }
    }

    function updateStats() {
        if (elStats) {
            elStats.textContent = `成功 ${successCount}  ·  失败 ${failCount}  ·  跳过 ${skipCount}`;
        }
        updateProgress();
        updateMiniButton();
    }

    function updateMiniButton() {
        if (!elMiniBtn) return;
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

    /** 浏览器最小化/后台时保持脚本活跃 */
    function startKeepAlive() {
        if (keepAliveTimer) return;
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
        keepAliveTimer = setInterval(() => {
            if (running && document.hidden) updateMiniButton();
        }, 900);
    }

    function stopKeepAlive() {
        if (keepAliveTimer) {
            clearInterval(keepAliveTimer);
            keepAliveTimer = null;
        }
    }

    function getVisibleText() {
        return (document.body && document.body.innerText) || '';
    }

    function detectRateLimit() {
        const text = getVisibleText();
        return (
            text.includes('请不要频繁点击') ||
            text.includes('1分钟后再试') ||
            text.includes('请1分钟后再试') ||
            text.includes('请一分钟后再试')
        );
    }

    function detectUserNotFound() {
        const text = getVisibleText();
        const hints = [
            '用户不存在', '该用户不存在', '未找到该用户', '找不到用户',
            '用户未找到', '无此用户', '未搜索到', '搜索无结果', '不存在该用户'
        ];
        return hints.some(h => text.includes(h));
    }

    async function waitRateLimitCooldown(phone) {
        await dismissOverlaySafely();
        const range = getDelayRangeMinutes();
        const totalMs = randomDelayMs(range.min, range.max);
        const endAt = Date.now() + totalMs;
        while (Date.now() < endAt) {
            if (stopRequested) return;
            const leftSec = Math.ceil((endAt - Date.now()) / 1000);
            const minPart = Math.floor(leftSec / 60);
            const secPart = leftSec % 60;
            const timeText = minPart > 0 ? `${minPart}分${String(secPart).padStart(2, '0')}秒` : `${secPart}秒`;
            setStatus(`频率限制，随机等待 ${timeText} 后重试（${range.min}-${range.max} 分钟）: ${phone}`);
            await delay(1000);
        }
    }

    function parseDelayMinutes(raw, fallback) {
        const s = String(raw ?? '').trim();
        if (!s) return fallback;
        const n = parseFloat(s);
        return Number.isFinite(n) && n >= 0 ? n : fallback;
    }

    function getDelayRangeMinutes() {
        let minM = parseDelayMinutes(elDelayMin && elDelayMin.value, DEFAULT_DELAY_MIN);
        let maxM = parseDelayMinutes(elDelayMax && elDelayMax.value, DEFAULT_DELAY_MAX);
        if (minM > maxM) [minM, maxM] = [maxM, minM];
        return { min: minM, max: maxM };
    }

    function randomDelayMs(minM, maxM) {
        const minMs = minM * 60 * 1000;
        const maxMs = maxM * 60 * 1000;
        if (minMs >= maxMs) return minMs;
        return minMs + Math.random() * (maxMs - minMs);
    }

    function saveDelaySettings() {
        try {
            const range = getDelayRangeMinutes();
            localStorage.setItem(DELAY_STORAGE_KEY, JSON.stringify(range));
        } catch (e) {}
    }

    function loadDelaySettings() {
        try {
            const raw = localStorage.getItem(DELAY_STORAGE_KEY);
            if (!raw) return;
            const data = JSON.parse(raw);
            if (elDelayMin && data.min != null) elDelayMin.value = data.min;
            if (elDelayMax && data.max != null) elDelayMax.value = data.max;
        } catch (e) {}
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

    function saveProgress() {
        try {
            if (!phoneList.length) return;
            localStorage.setItem(PROGRESS_STORAGE_KEY, JSON.stringify({
                phoneList,
                phoneIndex,
                successCount,
                failCount,
                skipCount,
                talkIndex,
                fingerprint: listFingerprint(phoneList),
            }));
        } catch (e) {}
    }

    function loadProgress() {
        try {
            const raw = localStorage.getItem(PROGRESS_STORAGE_KEY);
            if (!raw) return false;
            const data = JSON.parse(raw);
            if (!data || !Array.isArray(data.phoneList) || !data.phoneList.length) return false;
            if (data.fingerprint !== listFingerprint(data.phoneList)) return false;

            phoneList = data.phoneList;
            phoneIndex = Math.min(Math.max(0, data.phoneIndex || 0), phoneList.length);
            successCount = data.successCount || 0;
            failCount = data.failCount || 0;
            skipCount = data.skipCount || 0;
            talkIndex = data.talkIndex || 0;
            updateStats();
            updateStartButtonLabel();
            return true;
        } catch (e) {
            return false;
        }
    }

    function clearProgress() {
        try {
            localStorage.removeItem(PROGRESS_STORAGE_KEY);
        } catch (e) {}
        updateStartButtonLabel();
    }

    function updateStartButtonLabel() {
        if (!elBtnStart || running) return;
        const hasResume = phoneList.length > 0 && phoneIndex > 0 && phoneIndex < phoneList.length;
        elBtnStart.textContent = hasResume ? '继续' : '开始';
    }

    async function simulateTyping(input, value) {
        input.focus();
        const nativeSetter = Object.getOwnPropertyDescriptor(input.__proto__, 'value').set;
        nativeSetter.call(input, '');
        nativeSetter.call(input, value);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        await delay(120);
        ['keydown', 'keypress', 'keyup'].forEach(type => {
            input.dispatchEvent(new KeyboardEvent(type, {
                key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
            }));
        });
        await delay(150);
    }

    function waitForSelector(selector, retry = 40, interval = 200) {
        return new Promise(resolve => {
            let count = 0;
            const timer = setInterval(() => {
                const el = document.querySelector(selector);
                if (el) {
                    clearInterval(timer);
                    resolve(el);
                } else if (++count >= retry) {
                    clearInterval(timer);
                    resolve(null);
                }
            }, interval);
        });
    }

    function waitForButton(keyword, role = 'button', retry = 60, interval = 250, root = document) {
        return new Promise(resolve => {
            let count = 0;
            const timer = setInterval(() => {
                const btn = Array.from(root.querySelectorAll(`button, div[role="${role}"]`))
                    .find(el => {
                        if (isOurUI(el)) return false;
                        return (el.innerText || '').includes(keyword);
                    });
                if (btn) {
                    clearInterval(timer);
                    resolve(btn);
                } else if (++count >= retry) {
                    clearInterval(timer);
                    resolve(null);
                }
            }, interval);
        });
    }

    function findExistingFriendPanel() {
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

    async function closeExistingFriendPanel(panel) {
        panel = panel || findExistingFriendPanel();
        if (!panel) return false;

        const root = panel.closest('.semi-modal, .modal, [class*="modal"]') || panel;
        const box = root.getBoundingClientRect();

        const scopedClose = root.querySelector(
            '.semi-modal-close, .semi-icons-close, [class*="modal-close"], [class*="Modal-close"]'
        );
        if (scopedClose) {
            scopedClose.click();
            await delay(450);
            return !findExistingFriendPanel();
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
            await delay(450);
            return !findExistingFriendPanel();
        }

        return false;
    }

    async function waitForExistingFriendPanel(retry = 24, interval = 250) {
        return new Promise(resolve => {
            let count = 0;
            const timer = setInterval(() => {
                const p = findExistingFriendPanel();
                if (p) {
                    clearInterval(timer);
                    resolve(p);
                } else if (++count >= retry) {
                    clearInterval(timer);
                    resolve(null);
                }
            }, interval);
        });
    }

    async function dismissOverlaySafely() {
        const friendPanel = findExistingFriendPanel();
        if (friendPanel) return closeExistingFriendPanel(friendPanel);
        return clickConfirmComplete();
    }

    /** 已是好友：发送消息+设置备注等面板 → 点 X 关闭 → 跳过当前号 */
    async function handleAlreadyFriend(phone, waitPanel) {
        let panel = findExistingFriendPanel();
        if (!panel && waitPanel) panel = await waitForExistingFriendPanel(12, 200);
        if (!panel) return null;

        await closeExistingFriendPanel(panel);
        await delay(400);
        if (findExistingFriendPanel()) {
            await closeExistingFriendPanel();
            await delay(400);
        }
        if (findExistingFriendPanel()) {
            const back = document.querySelector(
                '.wk-chat-conversation-header-back, [class*="header-back"], [class*="nav-back"]'
            );
            if (back && !isOurUI(back)) {
                back.click();
                await delay(450);
            }
        }
        setStatus('已是好友，已跳过: ' + phone);
        return 'already_friend';
    }

    /** 频率限制 / 用户不存在（不改变 phoneIndex 的仅 rate_limit、retry） */
    function checkRateLimitOrNotFound() {
        if (detectRateLimit()) return 'rate_limit';
        if (detectUserNotFound()) return 'not_found';
        return null;
    }

    async function checkAfterAction(phone, waitFriendPanel) {
        const blocked = checkRateLimitOrNotFound();
        if (blocked) return blocked;
        return await handleAlreadyFriend(phone, waitFriendPanel);
    }

    async function findRemarkInput() {
        const container = document.querySelector('.semi-modal-content, .modal-content, .popup-content') || document.body;
        const textarea = container.querySelector('textarea.semi-textarea, textarea');
        if (textarea) return textarea;

        const inputs = container.querySelectorAll('input[type="text"]');
        for (const inp of inputs) {
            const ph = (inp.getAttribute('placeholder') || '').toLowerCase();
            if (ph.includes('备注') || ph.includes('留言')) return inp;
        }

        return container.querySelector('div[contenteditable="true"]') || null;
    }

    async function checkRemarkFilled(input, expected) {
        await delay(200);
        if (input.tagName.toLowerCase() === 'div') return input.textContent.trim() === expected;
        return input.value.trim() === expected;
    }

    async function clickConfirmComplete() {
        try {
            const modal = document.querySelector('.semi-modal-content, .modal-content, .popup-content');
            const root = modal || document.body;

            const completeBtn = await waitForButton('完成', 'button', 20, 250, root);
            if (completeBtn && !completeBtn.disabled) {
                completeBtn.click();
                await delay(400);
                return true;
            }
            const okBtn = await waitForButton('确定', 'button', 20, 250, root);
            if (okBtn && !okBtn.disabled) {
                okBtn.click();
                await delay(400);
                return true;
            }
            const confirmBtn = await waitForButton('确认', 'button', 20, 250, root);
            if (confirmBtn && !confirmBtn.disabled) {
                confirmBtn.click();
                await delay(400);
                return true;
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
     * rate_limit   → 频繁限制，随机等待后重试当前号
     * retry        → 未完成添加，清弹窗后重试当前号
     */
    async function addFriendAttempt(phone) {
        if (detectRateLimit()) return 'rate_limit';

        const input = await waitForSelector(
            'input.semi-input, input[type="text"], input[placeholder*="手机号"], input[placeholder*="畅言"]'
        );
        if (!input) return 'retry';

        await simulateTyping(input, phone);
        await delay(600);

        let outcome = await checkAfterAction(phone, true);
        if (outcome) return outcome;

        const addBtn = await waitForButton('添加好友', 'button', 40, 250);
        if (!addBtn) {
            outcome = await checkAfterAction(phone, true);
            if (outcome) return outcome;
            return 'retry';
        }

        addBtn.click();
        await delay(600);

        outcome = await checkAfterAction(phone, false);
        if (outcome) return outcome;

        const remarkInput = await findRemarkInput();
        if (!remarkInput) {
            outcome = await checkAfterAction(phone, true);
            if (outcome) return outcome;
            return 'retry';
        }
        if (talkList.length === 0) return 'retry';

        const msg = talkList[talkIndex++ % talkList.length];
        if (remarkInput.tagName.toLowerCase() === 'div') {
            remarkInput.focus();
            document.execCommand('selectAll');
            document.execCommand('delete');
            remarkInput.textContent = msg;
            remarkInput.dispatchEvent(new InputEvent('input', { bubbles: true }));
        } else {
            await simulateTyping(remarkInput, msg);
        }

        const ok = await checkRemarkFilled(remarkInput, msg);
        if (!ok) return 'retry';

        await clickConfirmComplete();
        await delay(500);

        outcome = checkRateLimitOrNotFound();
        if (outcome) return outcome;

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
            clearProgress();
        }

        running = true;
        stopRequested = false;
        updateStats();
        startKeepAlive();
        elBtnStart.disabled = true;
        elBtnStop.disabled = false;
        elBtnImport.disabled = true;
        if (elBtnClear) elBtnClear.disabled = true;
        if (elDelayMin) elDelayMin.disabled = true;
        if (elDelayMax) elDelayMax.disabled = true;
        if (elTalk) elTalk.disabled = true;

        if (phoneIndex > 0) {
            setStatus(`继续执行，从第 ${phoneIndex + 1}/${phoneList.length} 个开始…`);
        } else {
            setStatus('开始加好友（后台可继续）...');
        }

        while (phoneIndex < phoneList.length) {
            if (stopRequested) {
                saveProgress();
                setStatus(`已暂停，进度已保存 | 下次从第 ${phoneIndex + 1}/${phoneList.length} 个继续 | 成功 ${successCount} · 失败 ${failCount} · 跳过 ${skipCount}`);
                break;
            }

            const phone = phoneList[phoneIndex];
            updateStats();
            setStatus(`处理: ${phone}`);

            const result = await addFriendAttempt(phone);

            if (result === 'success') {
                successCount++;
                phoneIndex++;
                updateStats();
                saveProgress();
                setStatus(`成功 ${successCount} · 失败 ${failCount} · 跳过 ${skipCount} | 已完成: ${phone}`);
                await delay(1200);
            } else if (result === 'not_found') {
                failCount++;
                phoneIndex++;
                updateStats();
                saveProgress();
                setStatus(`用户不存在，跳过: ${phone} | 成功 ${successCount} · 失败 ${failCount} · 跳过 ${skipCount}`);
                await dismissOverlaySafely();
                await delay(1200);
            } else if (result === 'already_friend') {
                skipCount++;
                phoneIndex++;
                updateStats();
                saveProgress();
                setStatus(`已是好友，跳过: ${phone} | 成功 ${successCount} · 失败 ${failCount} · 跳过 ${skipCount}`);
                await delay(1000);
            } else if (result === 'rate_limit') {
                await waitRateLimitCooldown(phone);
                await dismissOverlaySafely();
                saveProgress();
            } else {
                setStatus(`未完成添加，重试当前号: ${phone}`);
                await dismissOverlaySafely();
                await delay(1500);
            }

            if (stopRequested) {
                saveProgress();
                setStatus(`已暂停，进度已保存 | 下次从第 ${phoneIndex + 1}/${phoneList.length} 个继续 | 成功 ${successCount} · 失败 ${failCount} · 跳过 ${skipCount}`);
                break;
            }
        }

        if (!stopRequested && phoneIndex >= phoneList.length) {
            setStatus(`全部完成 | 成功 ${successCount} · 失败 ${failCount} · 跳过 ${skipCount}`);
            clearProgress();
        }

        running = false;
        stopKeepAlive();
        elBtnStart.disabled = false;
        elBtnStop.disabled = true;
        elBtnImport.disabled = false;
        if (elBtnClear) elBtnClear.disabled = false;
        if (elDelayMin) elDelayMin.disabled = false;
        if (elDelayMax) elDelayMax.disabled = false;
        if (elTalk) elTalk.disabled = false;
        updateStartButtonLabel();
        updateMiniButton();
    }

    function stopLoop() {
        stopRequested = true;
        setStatus('正在暂停并保存进度...');
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
        const skipCols = new Set();

        headerCells.forEach((cell, idx) => {
            const kind = classifyColumnHeader(cell);
            if (kind === 'skip') skipCols.add(idx);
            else if (kind === 'phone' || kind === 'changyan') {
                targetCols.push(idx);
                colKinds[idx] = kind;
            }
        });

        return { targetCols, colKinds, skipCols };
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
                skipCols: built.skipCols,
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
                plan = { headerIdx: -1, targetCols: [], colKinds: {}, skipCols: new Set(), mode: 'single', singleCol: bestCol };
            } else {
                plan = { headerIdx: -1, targetCols: [], colKinds: {}, skipCols: new Set(), mode: 'multi', singleCol: -1 };
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
                clearProgress();
                saveProgress();
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
            #cy-add-friend-panel {
                position: fixed; top: 16px; right: 16px; z-index: 99999;
                width: 380px; background: #f8fafc; border-radius: 14px;
                box-shadow: 0 12px 40px rgba(15,23,42,0.16), 0 0 0 1px rgba(15,23,42,0.06);
                font-family: "Segoe UI", "Microsoft YaHei UI", "PingFang SC", system-ui, sans-serif;
                font-size: 13px; color: #0f172a; overflow: hidden;
            }
            #cy-add-friend-panel.cy-minimized { display: none !important; }
            #cy-add-friend-panel .cy-head {
                display: flex; align-items: center; justify-content: space-between;
                padding: 12px 14px; background: linear-gradient(135deg,#15803d,#16a34a);
                color: #fff; cursor: move; user-select: none;
            }
            #cy-add-friend-panel .cy-head-title { font-weight: 700; font-size: 15px; line-height: 1.25; }
            #cy-add-friend-panel .cy-head-sub { font-size: 11px; opacity: 0.92; margin-top: 3px; }
            #cy-add-friend-panel .cy-head-btn {
                background: rgba(255,255,255,0.22); border: none; color: #fff;
                width: 28px; height: 28px; border-radius: 8px; cursor: pointer;
                line-height: 1; font-size: 18px; flex-shrink: 0;
            }
            #cy-add-friend-panel .cy-head-btn:hover { background: rgba(255,255,255,0.32); }
            #cy-add-friend-panel .cy-body { padding: 14px; }
            #cy-add-friend-panel .cy-card {
                background: #fff; border: 1px solid #e2e8f0; border-radius: 10px;
                padding: 10px 12px; margin-bottom: 12px;
            }
            #cy-add-friend-panel .cy-progress {
                font-size: 13px; color: #1e293b; font-weight: 700; margin-bottom: 6px;
            }
            #cy-add-friend-panel .cy-stats {
                font-size: 12px; color: #15803d; font-weight: 600;
                padding: 6px 8px; background: #f0fdf4; border-radius: 8px;
                border: 1px solid #bbf7d0; text-align: center;
            }
            #cy-add-friend-panel .cy-progress-track {
                height: 6px; background: #e2e8f0; border-radius: 99px;
                overflow: hidden; margin-top: 8px;
            }
            #cy-add-friend-panel .cy-progress-bar {
                height: 100%; width: 0%; border-radius: 99px;
                background: linear-gradient(90deg,#22c55e,#16a34a);
                transition: width .35s ease;
            }
            #cy-add-friend-panel .cy-split-row {
                display: flex; gap: 10px; margin-bottom: 12px; align-items: stretch;
            }
            #cy-add-friend-panel .cy-talk-col { flex: 1.35; min-width: 0; }
            #cy-add-friend-panel .cy-delay-col {
                flex: 0.85; min-width: 0; background: #fff;
                border: 1px solid #e2e8f0; border-radius: 10px; padding: 10px;
            }
            #cy-add-friend-panel .cy-label {
                display: block; margin-bottom: 6px; color: #475569;
                font-size: 12px; font-weight: 600; line-height: 1.3;
            }
            #cy-add-friend-panel textarea {
                width: 100%; box-sizing: border-box; min-height: 108px; resize: vertical;
                border: 1px solid #cbd5e1; border-radius: 8px; padding: 9px 10px;
                font-size: 13px; line-height: 1.45; font-family: inherit;
                background: #fff; color: #0f172a;
            }
            #cy-add-friend-panel textarea:focus {
                outline: none; border-color: #22c55e; box-shadow: 0 0 0 3px rgba(34,197,94,0.15);
            }
            #cy-add-friend-panel .cy-delay-box {
                display: flex; flex-direction: column; gap: 8px; margin-top: 2px;
            }
            #cy-add-friend-panel .cy-delay-inline {
                display: flex; align-items: center; justify-content: center;
                gap: 6px; flex-wrap: wrap;
            }
            #cy-add-friend-panel .cy-delay-input {
                width: 48px; padding: 7px 4px; border: 1px solid #cbd5e1; border-radius: 8px;
                font-size: 14px; font-weight: 600; text-align: center; font-family: inherit;
                color: #0f172a; background: #f8fafc;
            }
            #cy-add-friend-panel .cy-delay-input:focus {
                outline: none; border-color: #22c55e; box-shadow: 0 0 0 3px rgba(34,197,94,0.12);
            }
            #cy-add-friend-panel .cy-delay-sep { color: #94a3b8; font-weight: 600; }
            #cy-add-friend-panel .cy-delay-unit { color: #64748b; font-size: 12px; }
            #cy-add-friend-panel .cy-delay-hint {
                font-size: 11px; color: #94a3b8; line-height: 1.45; text-align: center;
            }
            #cy-add-friend-panel .cy-btns-wrap { display: flex; flex-direction: column; gap: 8px; margin-bottom: 10px; }
            #cy-add-friend-panel .cy-btns-row { display: flex; gap: 8px; }
            #cy-add-friend-panel button.cy-btn {
                flex: 1; min-width: 0; padding: 10px 8px; border: none;
                border-radius: 8px; cursor: pointer; color: #fff; font-size: 13px;
                font-weight: 600; font-family: inherit; transition: filter .15s ease;
            }
            #cy-add-friend-panel button.cy-btn:hover:not(:disabled) { filter: brightness(1.06); }
            #cy-add-friend-panel button.cy-btn:disabled:not(.cy-btn-stop) {
                opacity: 0.55; cursor: not-allowed;
            }
            #cy-add-friend-panel .cy-btn-import { background: #2563eb; }
            #cy-add-friend-panel .cy-btn-clear { background: #64748b; }
            #cy-add-friend-panel .cy-btn-start { background: #16a34a; }
            #cy-add-friend-panel button.cy-btn.cy-btn-stop {
                background: #dc2626 !important; color: #fff !important;
            }
            #cy-add-friend-panel button.cy-btn.cy-btn-stop:disabled {
                background: #dc2626 !important; color: #fff !important; opacity: 0.5;
            }
            #cy-add-friend-panel .cy-status {
                padding: 10px 11px; background: #fff; border: 1px solid #e2e8f0;
                border-radius: 10px; font-size: 12px; color: #475569;
                min-height: 42px; word-break: break-all; line-height: 1.55;
            }
            #cy-mini-btn {
                position: fixed; left: 8px; bottom: 68px; z-index: 9999;
                width: 48px; height: 48px; border-radius: 50%; border: none;
                background: linear-gradient(135deg, #16a34a, #15803d);
                color: #fff; cursor: pointer;
                box-shadow: 0 4px 16px rgba(0,0,0,0.25);
                display: none; flex-direction: column; align-items: center; justify-content: center;
                user-select: none; padding: 0; line-height: 1.05;
            }
            #cy-mini-btn.cy-visible { display: flex; }
            #cy-mini-btn.cy-running { box-shadow: 0 0 0 3px rgba(22,163,74,0.35), 0 4px 16px rgba(0,0,0,0.25); }
            #cy-mini-btn .cy-mini-main { font-size: 12px; font-weight: 700; }
            #cy-mini-btn .cy-mini-sub { font-size: 8px; margin-top: 1px; opacity: 0.95; }
        `;
        document.head.appendChild(style);

        panel = document.createElement('div');
        panel.id = 'cy-add-friend-panel';

        const head = document.createElement('div');
        head.className = 'cy-head';
        head.innerHTML = `
            <div>
                <div class="cy-head-title">畅言加好友</div>
                <div class="cy-head-sub">阿陌专用 · 后台稳定版9.7</div>
            </div>
            <button type="button" class="cy-head-btn" id="cy-panel-minimize" title="最小化">−</button>
        `;

        const body = document.createElement('div');
        body.className = 'cy-body';

        const progressCard = document.createElement('div');
        progressCard.className = 'cy-card';

        elProgress = document.createElement('div');
        elProgress.className = 'cy-progress';
        elProgress.textContent = '0 / 0';

        const progressTrack = document.createElement('div');
        progressTrack.className = 'cy-progress-track';
        elProgressBar = document.createElement('div');
        elProgressBar.className = 'cy-progress-bar';
        progressTrack.appendChild(elProgressBar);

        elStats = document.createElement('div');
        elStats.className = 'cy-stats';
        elStats.textContent = '成功 0  ·  失败 0  ·  跳过 0';
        updateStats();

        progressCard.append(elProgress, elStats, progressTrack);

        const splitRow = document.createElement('div');
        splitRow.className = 'cy-split-row';

        const talkCol = document.createElement('div');
        talkCol.className = 'cy-talk-col';
        const talkLabel = document.createElement('label');
        talkLabel.className = 'cy-label';
        talkLabel.textContent = '话术（每行一条）';
        elTalk = document.createElement('textarea');
        elTalk.placeholder = '你好，很高兴认识你\n方便加个好友吗';
        elTalk.value = loadTalks();
        elTalk.addEventListener('blur', saveTalks);
        talkCol.append(talkLabel, elTalk);

        const delayCol = document.createElement('div');
        delayCol.className = 'cy-delay-col';
        const delayLabel = document.createElement('label');
        delayLabel.className = 'cy-label';
        delayLabel.textContent = '随机延迟（分钟）';
        const delayBox = document.createElement('div');
        delayBox.className = 'cy-delay-box';
        const delayInline = document.createElement('div');
        delayInline.className = 'cy-delay-inline';
        elDelayMin = Object.assign(document.createElement('input'), {
            type: 'number', className: 'cy-delay-input', min: '0', step: '0.5',
            placeholder: String(DEFAULT_DELAY_MIN), title: '最小分钟，默认1'
        });
        const delaySep = document.createElement('span');
        delaySep.className = 'cy-delay-sep';
        delaySep.textContent = '～';
        elDelayMax = Object.assign(document.createElement('input'), {
            type: 'number', className: 'cy-delay-input', min: '0', step: '0.5',
            placeholder: String(DEFAULT_DELAY_MAX), title: '最大分钟，默认5'
        });
        const delayUnit = document.createElement('span');
        delayUnit.className = 'cy-delay-unit';
        delayUnit.textContent = '分钟';
        elDelayMin.addEventListener('blur', saveDelaySettings);
        elDelayMax.addEventListener('blur', saveDelaySettings);
        delayInline.append(elDelayMin, delaySep, elDelayMax, delayUnit);
        const delayHint = document.createElement('div');
        delayHint.className = 'cy-delay-hint';
        delayHint.textContent = '频繁时等待，默认 1～5 分钟';
        delayBox.append(delayInline, delayHint);
        delayCol.append(delayLabel, delayBox);
        loadDelaySettings();

        splitRow.append(talkCol, delayCol);

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

        body.append(progressCard, splitRow, btnsWrap, elStatus);
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

        const drag = { active: false, x: 0, y: 0, left: 0, top: 0 };
        head.addEventListener('mousedown', e => {
            if (e.target.closest('button')) return;
            drag.active = true;
            drag.x = e.clientX;
            drag.y = e.clientY;
            const rect = panel.getBoundingClientRect();
            drag.left = rect.left;
            drag.top = rect.top;
            panel.style.right = 'auto';
            panel.style.left = drag.left + 'px';
            panel.style.top = drag.top + 'px';
        });
        document.addEventListener('mousemove', e => {
            if (!drag.active) return;
            panel.style.left = drag.left + (e.clientX - drag.x) + 'px';
            panel.style.top = drag.top + (e.clientY - drag.y) + 'px';
        });
        document.addEventListener('mouseup', () => { drag.active = false; });

        setStatus('面板已就绪，可最小化后台运行');
    }

    function boot() {
        createPanel();
        if (loadProgress()) {
            setStatus(`已恢复进度，从第 ${phoneIndex + 1}/${phoneList.length} 个继续（点「继续」开始）`);
        }
        loadSheetLib().then(ok => {
            if (ok && !phoneList.length) setStatus('Excel 库已就绪，可导入文件');
            else if (!ok) setStatus('Excel 库加载失败，请刷新页面');
        });
    }

    if (document.readyState !== 'loading') boot();
    else window.addEventListener('DOMContentLoaded', boot);

    setInterval(() => {
        if (!document.getElementById('cy-add-friend-panel')) createPanel();
        if (running && !keepAliveTimer) startKeepAlive();
    }, 4000);

})();
