/**
 * 爱租机委外催收系统小助手 - 主脚本
 * 功能：批量查询订单信息（含预留联系人）、批量添加催记、批量查询还款状态
 * 密码验证通过 window.AiZuJiHelperInit(password) 入口
 */

(function () {
    'use strict';

    // ========== 全局配置 ==========
    const CONFIG = {
        version: '1.3.1',
        name: '爱租机小助手',
        logPrefix: '[爱租机小助手]',
        // 访问密码（远程校验，可随时改）
        password: '99999',
        // API基础地址
        apiBase: 'https://internet-backend-gateway.woaizuji.com/fundApplication',
        // 默认并发数
        concurrency: 2,
        // 随机延迟（毫秒），默认2-8秒
        randomDelay: {
            enabled: true,
            min: 2000,
            max: 8000,
        },
        // 轮询间隔（毫秒）
        pollInterval: 300,
        // 请求超时
        requestTimeout: 30000,
        // XLSX CDN地址
        xlsxCdn: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    };

    // ========== 日志系统 ==========
    const Log = {
        info(msg, ...args) {
            console.log(`%c${CONFIG.logPrefix} ${msg}`, 'color: #3b82f6;', ...args);
        },
        success(msg, ...args) {
            console.log(`%c${CONFIG.logPrefix} ✓ ${msg}`, 'color: #10b981; font-weight: bold;', ...args);
        },
        warn(msg, ...args) {
            console.warn(`%c${CONFIG.logPrefix} ⚠ ${msg}`, 'color: #f59e0b;', ...args);
        },
        error(msg, ...args) {
            console.error(`%c${CONFIG.logPrefix} ✗ ${msg}`, 'color: #ef4444; font-weight: bold;', ...args);
        },
        debug(msg, ...args) {
            console.debug(`%c${CONFIG.logPrefix} ${msg}`, 'color: #8b5cf6;', ...args);
        },
        group(label) {
            console.group(`%c${CONFIG.logPrefix} ${label}`, 'color: #6366f1; font-weight: bold;');
        },
        groupEnd() {
            console.groupEnd();
        },
    };

    // ========== 工具函数 ==========
    const Utils = {
        // 延迟
        sleep(ms) {
            return new Promise(resolve => setTimeout(resolve, ms));
        },

        // 确保XLSX库已加载
        _xlsxLoaded: null,
        ensureXLSX() {
            if (typeof XLSX !== 'undefined') {
                return Promise.resolve(XLSX);
            }
            if (this._xlsxLoaded) {
                return this._xlsxLoaded;
            }
            this._xlsxLoaded = new Promise((resolve, reject) => {
                const script = document.createElement('script');
                script.src = CONFIG.xlsxCdn;
                script.onload = () => {
                    Log.success('XLSX 库加载成功');
                    resolve(window.XLSX);
                };
                script.onerror = () => {
                    this._xlsxLoaded = null;
                    reject(new Error('XLSX 库加载失败'));
                };
                document.head.appendChild(script);
            });
            return this._xlsxLoaded;
        },

        // 并发控制的Promise池
        async asyncPool(poolLimit, array, iteratorFn) {
            const ret = [];
            const executing = [];
            for (const item of array) {
                const p = Promise.resolve().then(() => iteratorFn(item, array));
                ret.push(p);
                if (poolLimit <= array.length) {
                    const e = p.then(() => executing.splice(executing.indexOf(e), 1));
                    executing.push(e);
                    if (executing.length >= poolLimit) {
                        await Promise.race(executing);
                    }
                }
            }
            return Promise.all(ret);
        },

        // 读取Excel文件
        async readExcel(file) {
            await this.ensureXLSX();
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = function (e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
                        resolve(jsonData);
                    } catch (err) {
                        reject(err);
                    }
                };
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
        },

        // 导出Excel
        async exportExcel(data, filename) {
            await this.ensureXLSX();
            const ws = XLSX.utils.json_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, '查询结果');
            XLSX.writeFile(wb, filename);
        },

        // 格式化时间
        formatTime(date = new Date()) {
            const pad = n => String(n).padStart(2, '0');
            return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
        },

        // 从localStorage/cookie获取token
        getToken() {
            // 尝试从localStorage获取
            try {
                const keys = Object.keys(localStorage);
                for (const key of keys) {
                    const lowerKey = key.toLowerCase();
                    if (lowerKey.indexOf('token') !== -1 || lowerKey.indexOf('azjtk') !== -1) {
                        const val = localStorage.getItem(key);
                        if (val && val.length > 50) return val;
                    }
                }
            } catch (e) {}
            // 尝试从cookie获取
            try {
                const cookies = document.cookie.split(';');
                for (let i = 0; i < cookies.length; i++) {
                    const c = cookies[i].trim();
                    const eqIndex = c.indexOf('=');
                    if (eqIndex === -1) continue;
                    const k = c.substring(0, eqIndex);
                    const v = c.substring(eqIndex + 1);
                    const lowerK = k.toLowerCase();
                    if (lowerK.indexOf('token') !== -1 || lowerK.indexOf('azjtk') !== -1) {
                        return v;
                    }
                }
            } catch (e) {}
            return '';
        },

        // 发送POST请求（使用页面原生fetch，避免跨域问题）
        async post(url, body) {
            const token = this.getToken();
            const headers = {
                'accept': 'application/json, text/plain, */*',
                'accept-language': 'zh-CN,zh;q=0.9',
                'content-type': 'application/json',
            };
            if (token) {
                headers['azjtk'] = token;
            }

            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), CONFIG.requestTimeout);

            try {
                const response = await fetch(url, {
                    headers,
                    body: JSON.stringify(body),
                    method: 'POST',
                    mode: 'cors',
                    credentials: 'omit',
                    signal: controller.signal,
                });
                clearTimeout(timeoutId);
                return await response.json();
            } catch (err) {
                clearTimeout(timeoutId);
                throw err;
            }
        },

        // 发送GET请求
        async get(url) {
            const token = this.getToken();
            const headers = {
                'accept': 'application/json, text/plain, */*',
                'accept-language': 'zh-CN,zh;q=0.9',
            };
            if (token) {
                headers['azjtk'] = token;
            }

            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), CONFIG.requestTimeout);

            try {
                const response = await fetch(url, {
                    headers,
                    method: 'GET',
                    mode: 'cors',
                    credentials: 'omit',
                    signal: controller.signal,
                });
                clearTimeout(timeoutId);
                return await response.json();
            } catch (err) {
                clearTimeout(timeoutId);
                throw err;
            }
        },

        // 随机延迟（模拟人工操作）
        randomDelay(min = 2000, max = 8000) {
            const delay = Math.floor(Math.random() * (max - min + 1)) + min;
            return this.sleep(delay);
        },
    };

    // ========== API层 ==========
    const API = {
        // 查询订单详情
        async queryOrderDetail(orderSN) {
            const url = `${CONFIG.apiBase}/outOverdue/detail/order/queryOverdueOrderDetail`;
            const result = await Utils.post(url, { orderSN });
            return result;
        },

        // 查询预留联系人列表
        async queryContactList(orderSN) {
            const url = `${CONFIG.apiBase}/api/orderContact/queryIntegratedList`;
            const result = await Utils.post(url, {
                currPage: 1,
                pageSize: 20,
                totalNum: 0,
                totalPages: 0,
                orderSN,
                sourceType: [2],
            });
            return result;
        },

        // 查询订单列表（获取订单号列表）
        async queryOrderList(params = {}) {
            const url = `${CONFIG.apiBase}/outOverdue/order/outOverdueQueryPage`;
            const body = {
                pageSize: 1000,
                currPage: 1,
                operatorId: '',
                legalStatus: '',
                isAllocated: '-1',
                merchantCode: '',
                merchantType: '',
                merchantName: '',
                repairState: '',
                type: '',
                ...params,
            };
            const result = await Utils.post(url, body);
            return result;
        },

        // 添加催记
        async addContactRecord(data) {
            const url = `${CONFIG.apiBase}/outOverdue/detail/contactRecord/add`;
            const body = {
                userName: data.userName || '',
                userPhone: data.userPhone || '',
                contactResult: data.contactResult || '6',
                opPerson: data.opPerson || '',
                orderSN: data.orderSN || '',
                remark: data.remark || '无法接通',
            };
            const result = await Utils.post(url, body);
            return result;
        },

        // 获取外部用户列表
        async getOutUserList() {
            const url = `${CONFIG.apiBase}/user/outUserList`;
            const result = await Utils.get(url);
            return result;
        },

        // 查询还款交易明细（只取前3条）
        async getTransactionDetailPage(orderSN, pageSize = 50) {
            const url = `${CONFIG.apiBase}/outOverdue/detail/transactionDetailPage`;
            const body = {
                orderSN: orderSN,
                state: 2,
                currPage: 1,
                pageSize: pageSize,
            };
            const result = await Utils.post(url, body);
            return result;
        },
    };

    // ========== UI面板 ==========
    const UI = {
        panel: null,
        isMinimized: false,

        init() {
            this.createPanel();
            this.createToggleButton();
            Log.success('UI面板初始化完成');
        },

        createPanel() {
            const panel = document.createElement('div');
            panel.id = 'aizuji-helper-panel';
            panel.innerHTML = `
                <div class="azh-header">
                    <span class="azh-title">🛠️ 爱租机小助手 v${CONFIG.version}</span>
                    <div class="azh-header-actions">
                        <button class="azh-btn azh-btn-minimize" title="最小化">—</button>
                        <button class="azh-btn azh-btn-close" title="关闭">✕</button>
                    </div>
                </div>
                <div class="azh-body">
                    <div class="azh-tabs">
                        <button class="azh-tab active" data-tab="sms">批量查订单</button>
                        <button class="azh-tab" data-tab="collection">批量催记</button>
                        <button class="azh-tab" data-tab="repayment">批量查还款</button>
                        <button class="azh-tab" data-tab="settings">设置</button>
                    </div>
                    <div class="azh-tab-content">
                        <!-- 批量查询订单信息 -->
                        <div class="azh-tab-pane active" id="azh-tab-sms">
                            <div class="azh-section">
                                <h3>📋 批量查询订单信息</h3>
                                <p class="azh-desc">查询所有在库订单的详情+预留联系人，直接导出</p>
                                <div class="azh-form-group">
                                    <label>查询方式：</label>
                                    <div style="display:flex;gap:12px;margin-bottom:8px;">
                                        <label style="display:flex;align-items:center;gap:4px;font-weight:normal;cursor:pointer;">
                                            <input type="radio" name="sms-mode" value="all" checked>
                                            查询全部在库订单
                                        </label>
                                        <label style="display:flex;align-items:center;gap:4px;font-weight:normal;cursor:pointer;">
                                            <input type="radio" name="sms-mode" value="excel">
                                            上传Excel指定订单
                                        </label>
                                    </div>
                                </div>
                                <div class="azh-form-group" id="azh-sms-file-group" style="display:none;">
                                    <label>选择Excel文件：</label>
                                    <input type="file" id="azh-sms-file" accept=".xlsx,.xls" class="azh-file-input">
                                    <div class="azh-hint" style="margin-top:6px;font-size:11px;">
                                        第一列是订单号（列名：订单号 / orderSN 均可识别）
                                    </div>
                                </div>
                                <div class="azh-form-group">
                                    <label>并发数：</label>
                                    <input type="number" id="azh-sms-concurrency" value="2" min="1" max="20" class="azh-input-number">
                                    <span style="font-size:12px;color:#94a3b8;margin-left:8px;">建议5-10</span>
                                </div>
                                <div class="azh-form-group">
                                    <label>导出字段：</label>
                                    <div class="azh-hint">
                                        订单号、姓名、年龄、手机号、手机型号、逾期天数、应还总额、应还租金、应还滞纳金、已还租金、联系人姓名、联系电话
                                    </div>
                                </div>
                                <div class="azh-btn-group">
                                    <button id="azh-sms-start" class="azh-btn azh-btn-primary">🚀 开始查询</button>
                                    <button id="azh-sms-export" class="azh-btn azh-btn-secondary" style="display:none;">💾 导出结果</button>
                                </div>
                                <div id="azh-sms-progress" class="azh-progress-box" style="display:none;">
                                    <div class="azh-progress-bar"><div class="azh-progress-fill"></div></div>
                                    <div class="azh-progress-text">进度：0 / 0 （成功 0 / 失败 0）</div>
                                </div>
                                <div id="azh-sms-log" class="azh-log-box"></div>
                            </div>
                        </div>

                        <!-- 批量添加催记 -->
                        <div class="azh-tab-pane" id="azh-tab-collection">
                            <div class="azh-section">
                                <h3>📝 批量添加催记</h3>
                                <p class="azh-desc">上传Excel批量添加催记，contactResult默认6，remark默认无法接通</p>
                                <div class="azh-form-group">
                                    <label>选择Excel文件：</label>
                                    <input type="file" id="azh-collection-file" accept=".xlsx,.xls" class="azh-file-input">
                                    <div class="azh-hint" style="margin-top:6px;font-size:11px;">
                                        需包含：订单号、姓名、手机号（列名模糊匹配，操作人可留空使用下方选择）
                                        <a href="javascript:void(0)" id="azh-collection-download-template" style="color:#3b82f6;text-decoration:underline;margin-left:8px;">下载模板</a>
                                    </div>
                                </div>
                                <div class="azh-form-group">
                                    <label>操作人：<span style="font-weight:normal;font-size:11px;color:#94a3b8;margin-left:6px;">Excel未填写时使用此处选择的操作人</span></label>
                                    <select id="azh-collection-opPerson" class="azh-input-text" style="width:100%;">
                                        <option value="">加载中...</option>
                                    </select>
                                </div>
                                <div class="azh-form-group">
                                    <label>联系结果：<span style="font-weight:normal;font-size:11px;color:#94a3b8;margin-left:6px;">默认6（暂时无法联系），可自行修改</span></label>
                                    <input type="text" id="azh-collection-result" value="6" class="azh-input-text">
                                </div>
                                <div class="azh-form-group">
                                    <label>备注：<span style="font-weight:normal;font-size:11px;color:#94a3b8;margin-left:6px;">默认"无法接通"，可自行修改</span></label>
                                    <input type="text" id="azh-collection-remark" value="无法接通" class="azh-input-text" style="width:100%;">
                                </div>
                                <div class="azh-form-group">
                                    <label>并发数：</label>
                                    <input type="number" id="azh-collection-concurrency" value="2" min="1" max="20" class="azh-input-number">
                                    <span style="font-size:12px;color:#94a3b8;margin-left:8px;">建议3-5</span>
                                </div>
                                <div class="azh-btn-group">
                                    <button id="azh-collection-start" class="azh-btn azh-btn-primary">🚀 开始添加</button>
                                    <button id="azh-collection-export" class="azh-btn azh-btn-secondary" style="display:none;">💾 导出结果</button>
                                </div>
                                <div id="azh-collection-progress" class="azh-progress-box" style="display:none;">
                                    <div class="azh-progress-bar"><div class="azh-progress-fill"></div></div>
                                    <div class="azh-progress-text">进度：0 / 0 （成功 0 / 失败 0）</div>
                                </div>
                                <div id="azh-collection-log" class="azh-log-box"></div>
                            </div>
                        </div>

                        <!-- 批量查询还款状态 -->
                        <div class="azh-tab-pane" id="azh-tab-repayment">
                            <div class="azh-section">
                                <h3>💰 批量查询还款状态</h3>
                                <p class="azh-desc">批量查询订单还款交易明细，支持全部在库订单或Excel指定订单号</p>
                                <div class="azh-form-group">
                                    <label>查询方式：</label>
                                    <div style="display:flex;gap:16px;margin-top:6px;">
                                        <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">
                                            <input type="radio" name="repayment-mode" value="all" checked> 全部在库订单
                                        </label>
                                        <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">
                                            <input type="radio" name="repayment-mode" value="excel"> Excel指定订单号
                                        </label>
                                    </div>
                                </div>
                                <div class="azh-form-group" id="azh-repayment-file-group" style="display:none;">
                                    <label>选择Excel文件：</label>
                                    <input type="file" id="azh-repayment-file" accept=".xlsx,.xls" class="azh-file-input">
                                    <div class="azh-hint" style="margin-top:6px;font-size:11px;">需包含订单号列（列名模糊匹配）</div>
                                </div>
                                <div class="azh-hint" id="azh-repayment-month-hint" style="font-size:11px;color:#94a3b8;margin-bottom:8px;">
                                    💡 显示当月还款记录
                                </div>
                                <div class="azh-form-group">
                                    <label>并发数：</label>
                                    <input type="number" id="azh-repayment-concurrency" value="2" min="1" max="20" class="azh-input-number">
                                    <span style="font-size:12px;color:#94a3b8;margin-left:8px;">建议2-5</span>
                                </div>
                                <div class="azh-btn-group">
                                    <button id="azh-repayment-start" class="azh-btn azh-btn-primary">🚀 开始查询</button>
                                    <button id="azh-repayment-export" class="azh-btn azh-btn-secondary" style="display:none;">💾 导出结果</button>
                                </div>
                                <div id="azh-repayment-progress" class="azh-progress-box" style="display:none;">
                                    <div class="azh-progress-bar"><div class="azh-progress-fill"></div></div>
                                    <div class="azh-progress-text">进度：0 / 0</div>
                                </div>
                                <div id="azh-repayment-log" class="azh-log-box"></div>
                            </div>
                        </div>

                        <!-- 设置 -->
                        <div class="azh-tab-pane" id="azh-tab-settings">
                            <div class="azh-section">
                                <h3>⚙️ 设置</h3>
                                <div class="azh-form-group">
                                    <label>默认并发数：</label>
                                    <input type="number" id="azh-setting-concurrency" value="2" min="1" max="20" class="azh-input-number">
                                    <span style="font-size:12px;color:#94a3b8;margin-left:8px;">建议2-5</span>
                                </div>
                                <div class="azh-form-group">
                                    <label>
                                        <input type="checkbox" id="azh-setting-random-delay" checked>
                                        启用随机延迟（降低被检测风险）
                                    </label>
                                </div>
                                <div class="azh-form-group">
                                    <label>延迟范围（秒）：</label>
                                    <div style="display:flex;gap:8px;align-items:center;">
                                        <input type="number" id="azh-setting-delay-min" value="2" min="0" class="azh-input-number" style="width:80px;">
                                        <span>~</span>
                                        <input type="number" id="azh-setting-delay-max" value="8" min="0" class="azh-input-number" style="width:80px;">
                                    </div>
                                </div>
                                <button id="azh-save-settings" class="azh-btn azh-btn-primary"> 保存设置</button>
                                <button id="azh-reset-settings" class="azh-btn azh-btn-secondary" style="margin-left:8px;">🔄 恢复默认</button>
                            </div>
                        </div>
                    </div>
                </div>
            `;
            document.body.appendChild(panel);
            this.panel = panel;
            this.bindEvents();
            this.injectStyles();
            this.makeDraggable(panel.querySelector('.azh-header'), panel);
        },

        createToggleButton() {
            const btn = document.createElement('button');
            btn.id = 'aizuji-helper-toggle';
            btn.innerHTML = '🛠️';
            btn.title = '爱租机小助手';
            btn.style.display = 'none';
            btn.onclick = () => this.togglePanel();
            document.body.appendChild(btn);
        },

        bindEvents() {
            // 最小化/关闭
            this.panel.querySelector('.azh-btn-minimize').onclick = () => this.togglePanel();
            this.panel.querySelector('.azh-btn-close').onclick = () => this.hidePanel();

            // Tab切换
            this.panel.querySelectorAll('.azh-tab').forEach(tab => {
                tab.onclick = () => {
                    this.panel.querySelectorAll('.azh-tab').forEach(t => t.classList.remove('active'));
                    this.panel.querySelectorAll('.azh-tab-pane').forEach(p => p.classList.remove('active'));
                    tab.classList.add('active');
                    this.panel.querySelector(`#azh-tab-${tab.dataset.tab}`).classList.add('active');
                };
            });
        },

        togglePanel() {
            if (this.isMinimized) {
                this.panel.style.display = 'flex';
                document.getElementById('aizuji-helper-toggle').style.display = 'none';
                this.isMinimized = false;
            } else {
                this.panel.style.display = 'none';
                document.getElementById('aizuji-helper-toggle').style.display = 'block';
                this.isMinimized = true;
            }
        },

        hidePanel() {
            this.panel.style.display = 'none';
            document.getElementById('aizuji-helper-toggle').style.display = 'block';
            this.isMinimized = true;
        },

        makeDraggable(header, panel) {
            let isDragging = false;
            let startX, startY, startLeft, startTop;

            header.style.cursor = 'move';
            header.onmousedown = (e) => {
                if (e.target.tagName === 'BUTTON') return;
                isDragging = true;
                startX = e.clientX;
                startY = e.clientY;
                const rect = panel.getBoundingClientRect();
                startLeft = rect.left;
                startTop = rect.top;
                e.preventDefault();
            };

            document.addEventListener('mousemove', (e) => {
                if (!isDragging) return;
                const dx = e.clientX - startX;
                const dy = e.clientY - startY;
                panel.style.left = Math.max(0, startLeft + dx) + 'px';
                panel.style.top = Math.max(0, startTop + dy) + 'px';
                panel.style.right = 'auto';
            });

            document.addEventListener('mouseup', () => {
                isDragging = false;
            });
        },

        injectStyles() {
            const style = document.createElement('style');
            style.textContent = `
                #aizuji-helper-panel {
                    position: fixed;
                    top: 20px;
                    right: 20px;
                    width: 400px;
                    max-height: 85vh;
                    background: #ffffff;
                    border-radius: 12px;
                    box-shadow: 0 10px 40px rgba(0,0,0,0.15);
                    z-index: 2147483647;
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    font-size: 13px;
                    color: #1e293b;
                    overflow: hidden;
                    display: flex;
                    flex-direction: column;
                }
                .azh-header {
                    padding: 10px 14px;
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    color: white;
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    user-select: none;
                    flex-shrink: 0;
                }
                .azh-title { font-weight: 600; font-size: 14px; }
                .azh-header-actions { display: flex; gap: 6px; }
                .azh-btn {
                    padding: 4px 10px;
                    border: none;
                    border-radius: 6px;
                    cursor: pointer;
                    font-size: 12px;
                    transition: all 0.2s;
                }
                .azh-btn:hover { opacity: 0.9; transform: translateY(-1px); }
                .azh-btn-minimize, .azh-btn-close {
                    background: rgba(255,255,255,0.2);
                    color: white;
                    width: 24px;
                    height: 24px;
                    padding: 0;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                }
                .azh-btn-primary {
                    background: #667eea;
                    color: white;
                    padding: 8px 16px;
                    font-weight: 500;
                }
                .azh-btn-secondary {
                    background: #e2e8f0;
                    color: #475569;
                    padding: 8px 16px;
                    font-weight: 500;
                }
                .azh-btn:disabled {
                    opacity: 0.5;
                    cursor: not-allowed;
                    transform: none;
                }
                .azh-btn-group {
                    display: flex;
                    gap: 8px;
                }
                .azh-body {
                    flex: 1;
                    overflow-y: auto;
                    padding: 0;
                }
                .azh-tabs {
                    display: flex;
                    border-bottom: 1px solid #e2e8f0;
                    padding: 0 10px;
                    flex-shrink: 0;
                    background: #f8fafc;
                }
                .azh-tab {
                    padding: 10px 12px;
                    border: none;
                    background: none;
                    cursor: pointer;
                    font-size: 12px;
                    color: #64748b;
                    border-bottom: 2px solid transparent;
                    margin-bottom: -1px;
                    transition: all 0.2s;
                }
                .azh-tab:hover { color: #475569; }
                .azh-tab.active {
                    color: #667eea;
                    border-bottom-color: #667eea;
                    font-weight: 500;
                    background: white;
                }
                .azh-tab-pane { display: none; }
                .azh-tab-pane.active { display: block; }
                .azh-section {
                    padding: 14px;
                }
                .azh-section h3 {
                    margin: 0 0 6px 0;
                    font-size: 14px;
                    color: #1e293b;
                }
                .azh-desc {
                    margin: 0 0 12px 0;
                    font-size: 12px;
                    color: #64748b;
                }
                .azh-form-group {
                    margin-bottom: 12px;
                }
                .azh-form-group label {
                    display: block;
                    margin-bottom: 6px;
                    font-weight: 500;
                    font-size: 12px;
                    color: #334155;
                }
                .azh-file-input {
                    width: 100%;
                    padding: 6px;
                    border: 1px solid #e2e8f0;
                    border-radius: 6px;
                    font-size: 12px;
                }
                .azh-input-number {
                    width: 60px;
                    padding: 4px 8px;
                    border: 1px solid #e2e8f0;
                    border-radius: 6px;
                    font-size: 12px;
                }
                .azh-input-text {
                    padding: 6px 10px;
                    border: 1px solid #e2e8f0;
                    border-radius: 6px;
                    font-size: 12px;
                    width: 120px;
                }
                .azh-hint {
                    padding: 8px 10px;
                    background: #f8fafc;
                    border-radius: 6px;
                    font-size: 12px;
                    color: #475569;
                    line-height: 1.6;
                }
                .azh-progress-box {
                    margin-top: 12px;
                }
                .azh-progress-bar {
                    width: 100%;
                    height: 8px;
                    background: #e2e8f0;
                    border-radius: 4px;
                    overflow: hidden;
                }
                .azh-progress-fill {
                    height: 100%;
                    background: linear-gradient(90deg, #667eea, #764ba2);
                    border-radius: 4px;
                    width: 0%;
                    transition: width 0.3s;
                }
                .azh-progress-text {
                    margin-top: 6px;
                    font-size: 12px;
                    color: #64748b;
                    text-align: center;
                }
                .azh-log-box {
                    margin-top: 12px;
                    max-height: 180px;
                    overflow-y: auto;
                    padding: 8px;
                    background: #0f172a;
                    border-radius: 6px;
                    font-family: 'Consolas', 'Monaco', monospace;
                    font-size: 11px;
                    line-height: 1.6;
                    color: #94a3b8;
                    display: none;
                }
                .azh-log-box.show { display: block; }
                .azh-log-box .log-success { color: #4ade80; }
                .azh-log-box .log-error { color: #f87171; }
                .azh-log-box .log-warn { color: #fbbf24; }
                .azh-log-box .log-info { color: #60a5fa; }

                #aizuji-helper-toggle {
                    position: fixed;
                    bottom: 20px;
                    right: 20px;
                    width: 48px;
                    height: 48px;
                    border-radius: 50%;
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    color: white;
                    border: none;
                    font-size: 20px;
                    cursor: pointer;
                    box-shadow: 0 4px 20px rgba(102, 126, 234, 0.4);
                    z-index: 2147483647;
                    transition: transform 0.2s;
                }
                #aizuji-helper-toggle:hover {
                    transform: scale(1.1);
                }
            `;
            document.head.appendChild(style);
        },

        // 追加日志
        appendLog(tabId, message, type = 'info') {
            const logBox = document.getElementById(`azh-${tabId}-log`);
            if (!logBox) return;
            logBox.classList.add('show');
            const line = document.createElement('div');
            line.className = `log-${type}`;
            line.textContent = `[${Utils.formatTime()}] ${message}`;
            logBox.appendChild(line);
            logBox.scrollTop = logBox.scrollHeight;
        },

        // 清空日志
        clearLog(tabId) {
            const logBox = document.getElementById(`azh-${tabId}-log`);
            if (!logBox) return;
            logBox.innerHTML = '';
        },

        // 更新进度条
        updateProgress(tabId, current, total, success = 0, failed = 0) {
            const progressBox = document.getElementById(`azh-${tabId}-progress`);
            const fill = progressBox.querySelector('.azh-progress-fill');
            const text = progressBox.querySelector('.azh-progress-text');
            progressBox.style.display = 'block';
            const percent = total > 0 ? (current / total * 100) : 0;
            fill.style.width = percent + '%';
            text.textContent = `进度：${current} / ${total} （成功 ${success} / 失败 ${failed}）`;
        },
    };

    // ========== 模块1：批量查询订单信息 ==========
    const SmsModule = {
        results: [],
        isRunning: false,
        successCount: 0,
        failCount: 0,

        async start() {
            if (this.isRunning) return;

            const mode = document.querySelector('input[name="sms-mode"]:checked')?.value || 'all';
            const concurrency = parseInt(document.getElementById('azh-sms-concurrency').value) || CONFIG.concurrency;

            this.isRunning = true;
            this.results = [];
            this.successCount = 0;
            this.failCount = 0;
            document.getElementById('azh-sms-start').disabled = true;
            document.getElementById('azh-sms-export').style.display = 'none';

            try {
                Log.group('批量查询订单信息');
                let orderSNs = [];

                if (mode === 'all') {
                    // ===== 模式1：查询全部在库订单 =====
                    Log.info('开始获取全部在库订单...');
                    UI.appendLog('sms', '📡 正在获取全部在库订单列表...', 'info');

                    orderSNs = await this.getAllOrderSNs();
                    Log.info(`获取到 ${orderSNs.length} 个在库订单`);
                    UI.appendLog('sms', `✅ 共获取到 ${orderSNs.length} 个在库订单，开始查询详情...`, 'info');

                } else {
                    // ===== 模式2：上传Excel指定订单 =====
                    const fileInput = document.getElementById('azh-sms-file');
                    if (!fileInput.files.length) {
                        alert('请先选择Excel文件！');
                        return;
                    }

                    Log.info('开始解析Excel...');
                    UI.appendLog('sms', '📖 开始解析Excel文件...', 'info');

                    const data = await Utils.readExcel(fileInput.files[0]);
                    Log.info(`共读取 ${data.length} 条数据`);
                    UI.appendLog('sms', `✅ 读取到 ${data.length} 条数据`, 'info');

                    if (data.length === 0) {
                        alert('Excel文件为空！');
                        return;
                    }

                    orderSNs = this.extractOrderSNs(data);
                }

                if (orderSNs.length === 0) {
                    alert('没有可查询的订单号！');
                    return;
                }

                // 批量查询详情
                let completed = 0;
                const total = orderSNs.length;

                await Utils.asyncPool(concurrency, orderSNs, async (orderSN) => {
                    try {
                        const row = await this.queryOne(orderSN);
                        this.results.push(row);
                        this.successCount++;
                    } catch (err) {
                        Log.error(`查询 ${orderSN} 失败:`, err);
                        this.results.push({
                            '订单号': orderSN,
                            '查询状态': '失败',
                            '失败原因': err.message || '未知错误',
                        });
                        this.failCount++;
                    }
                    completed++;
                    UI.updateProgress('sms', completed, total, this.successCount, this.failCount);
                    if (completed % 20 === 0 || completed === total) {
                        UI.appendLog('sms', `⏳ 已完成 ${completed}/${total}（成功 ${this.successCount}，失败 ${this.failCount}）`, 'info');
                    }
                });

                Log.success(`查询完成，成功 ${this.successCount} 条，失败 ${this.failCount} 条`);
                UI.appendLog('sms', `🎉 查询完成！成功 ${this.successCount} 条，失败 ${this.failCount} 条`, 'success');
                document.getElementById('azh-sms-export').style.display = 'inline-block';

            } catch (err) {
                Log.error('批量查询失败:', err);
                UI.appendLog('sms', `❌ 操作失败: ${err.message}`, 'error');
                alert('操作失败：' + err.message);
            } finally {
                this.isRunning = false;
                document.getElementById('azh-sms-start').disabled = false;
                Log.groupEnd();
            }
        },

        // 获取全部在库订单的订单号
        async getAllOrderSNs() {
            const orderSNs = [];
            const pageSize = 1000;
            let currPage = 1;
            let totalPages = 1;

            while (currPage <= totalPages) {
                const resp = await API.queryOrderList({
                    pageSize,
                    currPage,
                    isAllocated: '-1',
                });

                if (resp && resp.data && resp.data.data) {
                    const list = resp.data.data;
                    for (const item of list) {
                        if (item.orderSN) {
                            orderSNs.push(item.orderSN);
                        }
                    }
                    totalPages = resp.data.totalPages || 1;
                    UI.appendLog('sms', `   获取第 ${currPage}/${totalPages} 页，${list.length} 条`, 'info');
                } else {
                    break;
                }

                currPage++;
                // 翻页加个小延迟
                await Utils.sleep(200);
            }

            return orderSNs;
        },

        // 从Excel数据中提取订单号
        extractOrderSNs(data) {
            const orderSNs = [];
            const firstRow = data[0];
            const keys = Object.keys(firstRow);

            // 查找订单号列
            let orderKey = null;
            const orderKeywords = ['订单号', 'orderSN', 'ordersn', 'OrderSN', '订单编号', 'order_no', 'orderNo'];

            for (const key of keys) {
                if (orderKeywords.some(kw => key.toLowerCase().includes(kw.toLowerCase()))) {
                    orderKey = key;
                    break;
                }
            }

            // 如果没找到，默认用第一列
            if (!orderKey) {
                orderKey = keys[0];
                Log.warn(`未找到订单号列，默认使用第一列: ${orderKey}`);
            }

            for (const row of data) {
                const val = (row[orderKey] || '').toString().trim();
                if (val) {
                    orderSNs.push(val);
                }
            }

            return orderSNs;
        },

        // 查询单个订单的完整信息
        async queryOne(orderSN) {
            // 1. 查询订单详情
            const detailResp = await API.queryOrderDetail(orderSN);
            const detail = (detailResp && detailResp.data) || {};

            // 2. 查询预留联系人
            let contacts = [];
            try {
                const contactResp = await API.queryContactList(orderSN);
                if (contactResp && contactResp.data && contactResp.data.data) {
                    contacts = contactResp.data.data;
                }
            } catch (e) {
                Log.warn(`查询联系人失败 ${orderSN}:`, e);
            }

            // 随机延迟（如果设置了）
            if (CONFIG.randomDelay.enabled) {
                await Utils.randomDelay(CONFIG.randomDelay.min, CONFIG.randomDelay.max);
            }

            // 3. 组装结果
            const contactNames = contacts.map(c => c.relationshipDesc || '').filter(Boolean).join(' / ');
            const contactPhones = contacts.map(c => c.contactPhone || '').filter(Boolean).join(' / ');

            return {
                '订单号': detail.orderSN || orderSN,
                '姓名': detail.userName || '',
                '年龄': detail.age || '',
                '手机号': detail.phone || '',
                '手机型号': detail.materielModelName || '',
                '逾期天数': detail.overdueDay || '',
                '应还总额': detail.dueAmount || '',
                '应还租金': detail.currentBalancez || '',
                '应还滞纳金': detail.lateFeeOut || '',
                '已还租金': detail.totalRentPaid || '',
                '联系人姓名': contactNames,
                '联系人电话': contactPhones,
                '查询状态': '成功',
            };
        },

        async export() {
            if (this.results.length === 0) {
                alert('没有可导出的数据！');
                return;
            }
            // 导出时去掉内部用的查询状态列（如果用户只看成功数据的话保留）
            const exportData = this.results.map(row => {
                const copy = { ...row };
                return copy;
            });
            await Utils.exportExcel(exportData, `爱租机订单查询_${Utils.formatTime().replace(/[:\s-]/g, '')}.xlsx`);
            Log.success('查询结果已导出');
            UI.appendLog('sms', '💾 结果已导出Excel', 'success');
        },
    };

    // ========== 模块2：批量添加催记 ==========
    const CollectionModule = {
        results: [],
        isRunning: false,
        successCount: 0,
        failCount: 0,

        async start() {
            if (this.isRunning) return;

            const fileInput = document.getElementById('azh-collection-file');
            if (!fileInput.files.length) {
                alert('请先选择Excel文件！');
                return;
            }

            const concurrency = parseInt(document.getElementById('azh-collection-concurrency').value) || CONFIG.concurrency;
            const contactResult = document.getElementById('azh-collection-result').value || '6';
            const remark = document.getElementById('azh-collection-remark').value || '无法接通';
            const defaultOpPerson = document.getElementById('azh-collection-opPerson').value || '';

            if (!defaultOpPerson) {
                alert('请选择操作人！');
                return;
            }

            this.isRunning = true;
            this.results = [];
            this.successCount = 0;
            this.failCount = 0;
            document.getElementById('azh-collection-start').disabled = true;
            document.getElementById('azh-collection-export').style.display = 'none';

            try {
                Log.group('批量添加催记');

                Log.info('开始解析Excel...');
                UI.appendLog('collection', '📖 开始解析Excel文件...', 'info');

                const data = await Utils.readExcel(fileInput.files[0]);
                Log.info(`共读取 ${data.length} 条数据`);
                UI.appendLog('collection', `✅ 读取到 ${data.length} 条数据`, 'info');

                if (data.length === 0) {
                    alert('Excel文件为空！');
                    return;
                }

                const records = this.parseRecords(data, contactResult, remark, defaultOpPerson);
                Log.info(`解析出 ${records.length} 条催记记录`);
                UI.appendLog('collection', `🔍 解析出 ${records.length} 条记录，开始批量添加...`, 'info');

                // 批量添加
                let completed = 0;
                const total = records.length;

                await Utils.asyncPool(concurrency, records, async (record) => {
                    try {
                        // 随机延迟
                        if (CONFIG.randomDelay.enabled) {
                            await Utils.randomDelay(CONFIG.randomDelay.min, CONFIG.randomDelay.max);
                        }
                        const resp = await API.addContactRecord(record);
                        const ok = resp && (resp.code === 200 || resp.code === 0 || resp.success === true || resp.result === true);
                        const msg = (resp && resp.message) || (resp && resp.msg) || '';

                        this.results.push({
                            '订单号': record.orderSN,
                            '姓名': record.userName,
                            '手机号': record.userPhone,
                            '操作人': record.opPerson,
                            '联系结果': record.contactResult,
                            '备注': record.remark,
                            '添加状态': ok ? '成功' : '失败',
                            '返回信息': msg || (ok ? '添加成功' : JSON.stringify(resp).substring(0, 100)),
                        });
                        if (ok) {
                            this.successCount++;
                        } else {
                            this.failCount++;
                        }
                    } catch (err) {
                        Log.error(`添加催记失败 ${record.orderSN}:`, err);
                        this.results.push({
                            '订单号': record.orderSN,
                            '姓名': record.userName,
                            '手机号': record.userPhone,
                            '操作人': record.opPerson,
                            '联系结果': record.contactResult,
                            '备注': record.remark,
                            '添加状态': '失败',
                            '返回信息': err.message || '网络错误',
                        });
                        this.failCount++;
                    }
                    completed++;
                    UI.updateProgress('collection', completed, total, this.successCount, this.failCount);
                    if (completed % 20 === 0 || completed === total) {
                        UI.appendLog('collection', `⏳ 已完成 ${completed}/${total}（成功 ${this.successCount}，失败 ${this.failCount}）`, 'info');
                    }
                });

                Log.success(`添加完成，成功 ${this.successCount} 条，失败 ${this.failCount} 条`);
                UI.appendLog('collection', `🎉 添加完成！成功 ${this.successCount} 条，失败 ${this.failCount} 条`, 'success');
                document.getElementById('azh-collection-export').style.display = 'inline-block';

            } catch (err) {
                Log.error('批量添加催记失败:', err);
                UI.appendLog('collection', `❌ 操作失败: ${err.message}`, 'error');
                alert('操作失败：' + err.message);
            } finally {
                this.isRunning = false;
                document.getElementById('azh-collection-start').disabled = false;
                Log.groupEnd();
            }
        },

        // 从Excel数据中提取催记字段
        parseRecords(data, contactResult, remark, defaultOpPerson) {
            const records = [];
            const firstRow = data[0];
            const keys = Object.keys(firstRow);

            // 模糊匹配列名
            const findKey = (keywords) => {
                for (const key of keys) {
                    const lowerKey = key.toLowerCase().trim();
                    for (const kw of keywords) {
                        if (lowerKey.indexOf(kw.toLowerCase()) !== -1) {
                            return key;
                        }
                    }
                }
                return null;
            };

            const orderKey = findKey(['订单号', 'orderSN', 'order_no', 'orderNo', '订单编号', '案件编号']);
            const nameKey = findKey(['姓名', 'name', 'userName', '客户姓名', '客户名']);
            const phoneKey = findKey(['手机号', '电话', 'phone', 'userPhone', '手机号码', '联系电话']);
            const opKey = findKey(['操作人', 'opPerson', '操作员', '催收员', '操作']);

            if (!orderKey) {
                alert('未找到订单号列，请检查Excel表头！');
                throw new Error('未找到订单号列');
            }

            for (const row of data) {
                const orderSN = (row[orderKey] || '').toString().trim();
                if (!orderSN) continue;

                const excelOpPerson = opKey ? (row[opKey] || '').toString().trim() : '';
                records.push({
                    orderSN: orderSN,
                    userName: nameKey ? (row[nameKey] || '').toString().trim() : '',
                    userPhone: phoneKey ? (row[phoneKey] || '').toString().trim() : '',
                    opPerson: excelOpPerson || defaultOpPerson || '',
                    contactResult: contactResult,
                    remark: remark,
                });
            }

            return records;
        },

        // 加载用户列表填充下拉框
        async loadUserList() {
            const select = document.getElementById('azh-collection-opPerson');
            if (!select) return;

            try {
                select.innerHTML = '<option value="">加载中...</option>';
                const result = await API.getOutUserList();
                if (result && result.success && Array.isArray(result.data)) {
                    const users = result.data;
                    select.innerHTML = '';
                    if (users.length === 0) {
                        select.innerHTML = '<option value="">无可用操作人</option>';
                        return;
                    }
                    users.forEach(user => {
                        const option = document.createElement('option');
                        option.value = user.realName || '';
                        option.textContent = user.realName || user.userName || '未知';
                        if (user.tenantName) {
                            option.textContent += `（${user.tenantName}）`;
                        }
                        select.appendChild(option);
                    });
                    Log.info(`操作人列表加载完成，共${users.length}人`);
                } else {
                    select.innerHTML = '<option value="">加载失败</option>';
                    Log.warn('操作人列表加载失败:', result?.errMsg || '未知错误');
                }
            } catch (err) {
                select.innerHTML = '<option value="">加载失败</option>';
                Log.error('加载操作人列表失败:', err);
            }
        },

        async export() {
            if (this.results.length === 0) {
                alert('没有可导出的数据！');
                return;
            }
            await Utils.exportExcel(this.results, `爱租机批量催记_${Utils.formatTime().replace(/[:\s-]/g, '')}.xlsx`);
            Log.success('催记结果已导出');
            UI.appendLog('collection', '💾 结果已导出Excel', 'success');
        },

        async downloadTemplate() {
            await Utils.ensureXLSX();
            const templateData = [
                { 订单号: 'SA2407060066031440', 姓名: '沈庆明', 手机号: '17605226750' },
            ];
            const ws = XLSX.utils.json_to_sheet(templateData);
            ws['!cols'] = [{ wch: 22 }, { wch: 12 }, { wch: 15 }];
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, '催记模板');
            XLSX.writeFile(wb, '爱租机批量催记模板.xlsx');
            Log.success('模板已下载');
        },
    };

    // ========== 模块3：批量查询还款状态 ==========
    const RepaymentModule = {
        results: [],
        isRunning: false,

        async start() {
            if (this.isRunning) return;

            const mode = document.querySelector('input[name="repayment-mode"]:checked').value;
            const concurrency = parseInt(document.getElementById('azh-repayment-concurrency').value) || CONFIG.concurrency;

            let orderSNs = [];

            this.isRunning = true;
            this.results = [];
            document.getElementById('azh-repayment-start').disabled = true;
            document.getElementById('azh-repayment-export').style.display = 'none';
            document.getElementById('azh-repayment-progress').style.display = 'block';
            UI.updateProgress('repayment', 0, 1, 0, 0);
            UI.clearLog('repayment');

            try {
                Log.group('批量查询还款状态');

                // 获取订单号列表
                if (mode === 'all') {
                    UI.appendLog('repayment', '📋 正在获取全部在库订单...', 'info');
                    orderSNs = await SmsModule.getAllOrderSNs();
                    UI.appendLog('repayment', `✅ 共获取到 ${orderSNs.length} 个订单`, 'success');
                } else {
                    const fileInput = document.getElementById('azh-repayment-file');
                    if (!fileInput.files.length) {
                        alert('请先选择Excel文件！');
                        this._cleanup();
                        return;
                    }
                    UI.appendLog('repayment', '📄 正在解析Excel...', 'info');
                    await Utils.ensureXLSX();
                    const data = await Utils.readExcel(fileInput.files[0]);
                    const firstRow = data[0];
                    const keys = Object.keys(firstRow);
                    const findKey = (keywords) => {
                        for (const key of keys) {
                            const lowerKey = key.toLowerCase().trim();
                            for (const kw of keywords) {
                                if (lowerKey.indexOf(kw.toLowerCase()) !== -1) return key;
                            }
                        }
                        return null;
                    };
                    const orderKey = findKey(['订单号', 'orderSN', 'order_no', 'orderNo', '订单编号', '案件编号']);
                    if (!orderKey) {
                        alert('未找到订单号列，请检查Excel表头！');
                        this._cleanup();
                        return;
                    }
                    orderSNs = data.map(row => (row[orderKey] || '').toString().trim()).filter(s => s);
                    UI.appendLog('repayment', `✅ 解析到 ${orderSNs.length} 个订单号`, 'success');
                }

                if (orderSNs.length === 0) {
                    alert('没有找到有效的订单号！');
                    this._cleanup();
                    return;
                }

                const total = orderSNs.length;
                let completed = 0;
                let hasRecordCount = 0;

                UI.appendLog('repayment', `🔍 开始查询 ${total} 个订单的还款记录，并发 ${concurrency}`, 'info');
                UI.updateProgress('repayment', 0, total, 0, 0);

                await Utils.asyncPool(concurrency, orderSNs, async (orderSN) => {
                    try {
                        // 随机延迟
                        if (CONFIG.randomDelay.enabled) {
                            await Utils.randomDelay(CONFIG.randomDelay.min, CONFIG.randomDelay.max);
                        }

                        const resp = await API.getTransactionDetailPage(orderSN, 50);
                        let allRecords = [];
                        if (resp && resp.success && resp.data) {
                            const list = resp.data.data || resp.data.records || resp.data.list || [];
                            // 只保留当月的还款记录
                            const now = new Date();
                            const currentMonth = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
                            allRecords = list.filter(record => {
                                const time = record.finishTimeStr || '';
                                return time.indexOf(currentMonth) === 0;
                            }).slice(0, 50);
                        }

                        if (allRecords.length > 0) {
                            hasRecordCount++;
                            for (const record of allRecords) {
                                this.results.push({
                                    '订单号': record.orderSN || orderSN,
                                    '姓名': record.realName || '',
                                    '还款类型': record.transactionSourceDesc || '',
                                    '还款时间': record.finishTimeStr || '',
                                    '还款金额': record.transactionAmount || '',
                                });
                            }
                        }
                    } catch (err) {
                        Log.error(`查询还款失败 ${orderSN}:`, err);
                        this.results.push({
                            '订单号': orderSN,
                            '姓名': '',
                            '还款类型': '',
                            '还款时间': '',
                            '还款金额': '查询失败：' + (err.message || '网络错误'),
                        });
                    }
                    completed++;
                    UI.updateProgress('repayment', completed, total, hasRecordCount, 0);
                    if (completed % 20 === 0 || completed === total) {
                        UI.appendLog('repayment', `⏳ 已完成 ${completed}/${total}（${hasRecordCount} 个有还款记录）`, 'info');
                    }
                });

                Log.success(`查询完成，共 ${this.results.length} 条还款记录`);
                UI.appendLog('repayment', `🎉 查询完成！共 ${hasRecordCount} 个订单有 ${this.results.length} 条还款记录`, 'success');
                document.getElementById('azh-repayment-export').style.display = 'inline-block';

            } catch (err) {
                Log.error('批量查询还款失败:', err);
                UI.appendLog('repayment', `❌ 操作失败: ${err.message}`, 'error');
                alert('操作失败：' + err.message);
            } finally {
                this.isRunning = false;
                document.getElementById('azh-repayment-start').disabled = false;
                Log.groupEnd();
            }
        },

        _formatState(state) {
            const stateMap = {
                1: '处理中',
                2: '成功',
                3: '失败',
                4: '作废',
                5: '退票',
            };
            return stateMap[state] || (state !== undefined && state !== null ? state : '');
        },

        async export() {
            if (this.results.length === 0) {
                alert('没有可导出的数据！');
                return;
            }
            await Utils.exportExcel(this.results, `爱租机还款明细_${Utils.formatTime().replace(/[:\s-]/g, '')}.xlsx`);
            Log.success('还款明细已导出');
            UI.appendLog('repayment', '💾 结果已导出Excel', 'success');
        },

        _cleanup() {
            this.isRunning = false;
            document.getElementById('azh-repayment-start').disabled = false;
            document.getElementById('azh-repayment-progress').style.display = 'none';
        },
    };

    // ========== 主入口 ==========
    let initialized = false;

    function init() {
        if (initialized) return { success: true, message: '已初始化' };

        Log.info(`v${CONFIG.version} 启动中...`);

        // 初始化UI
        UI.init();

        // 更新还款模块的月份提示
        const monthHint = document.getElementById('azh-repayment-month-hint');
        if (monthHint) {
            const now = new Date();
            monthHint.innerHTML = '💡 显示 ' + now.getFullYear() + '年' + (now.getMonth() + 1) + '月 的还款记录';
        }

        // 绑定功能按钮
        document.getElementById('azh-sms-start').onclick = () => SmsModule.start();
        document.getElementById('azh-sms-export').onclick = async () => SmsModule.export();
        document.getElementById('azh-collection-start').onclick = () => CollectionModule.start();
        document.getElementById('azh-collection-export').onclick = async () => CollectionModule.export();
        document.getElementById('azh-collection-download-template').onclick = () => CollectionModule.downloadTemplate();
        // 异步加载操作人列表
        setTimeout(() => CollectionModule.loadUserList(), 500);
        const repaymentBtn = document.getElementById('azh-repayment-start');
        if (repaymentBtn) repaymentBtn.onclick = () => RepaymentModule.start();
        const repaymentExportBtn = document.getElementById('azh-repayment-export');
        if (repaymentExportBtn) repaymentExportBtn.onclick = async () => RepaymentModule.export();

        // 查询方式切换 - 批量查订单
        document.querySelectorAll('input[name="sms-mode"]').forEach(radio => {
            radio.onchange = () => {
                const fileGroup = document.getElementById('azh-sms-file-group');
                if (radio.value === 'excel') {
                    fileGroup.style.display = 'block';
                } else {
                    fileGroup.style.display = 'none';
                }
            };
        });

        // 查询方式切换 - 批量查还款
        document.querySelectorAll('input[name="repayment-mode"]').forEach(radio => {
            radio.onchange = () => {
                const fileGroup = document.getElementById('azh-repayment-file-group');
                if (radio.value === 'excel') {
                    fileGroup.style.display = 'block';
                } else {
                    fileGroup.style.display = 'none';
                }
            };
        });

        // 保存设置：写入 CONFIG 并同步到所有模块输入框
        document.getElementById('azh-save-settings').onclick = () => {
            const newConcurrency = parseInt(document.getElementById('azh-setting-concurrency').value) || 2;
            const delayEnabled = document.getElementById('azh-setting-random-delay').checked;
            const delayMinSec = parseInt(document.getElementById('azh-setting-delay-min').value) || 2;
            const delayMaxSec = parseInt(document.getElementById('azh-setting-delay-max').value) || 8;
            const delayMinMs = delayMinSec * 1000;
            const delayMaxMs = delayMaxSec * 1000;

            if (delayMinMs > delayMaxMs) {
                alert('延迟最小值不能大于最大值！');
                return;
            }

            // 写入全局配置
            CONFIG.concurrency = newConcurrency;
            CONFIG.randomDelay.enabled = delayEnabled;
            CONFIG.randomDelay.min = delayMinMs;
            CONFIG.randomDelay.max = delayMaxMs;

            // 同步到批量查订单模块
            const smsConcurrencyEl = document.getElementById('azh-sms-concurrency');
            if (smsConcurrencyEl) smsConcurrencyEl.value = newConcurrency;

            // 同步到批量催记模块
            const colConcurrencyEl = document.getElementById('azh-collection-concurrency');
            if (colConcurrencyEl) colConcurrencyEl.value = newConcurrency;

            // 同步到批量查还款模块
            const repayConcurrencyEl = document.getElementById('azh-repayment-concurrency');
            if (repayConcurrencyEl) repayConcurrencyEl.value = newConcurrency;

            Log.success(`设置已保存：并发=${newConcurrency}，延迟=${delayEnabled ? delayMinSec + '~' + delayMaxSec + '秒' : '关闭'}`);
            alert(`设置已生效！\n并发数：${newConcurrency}\n延迟：${delayEnabled ? delayMinSec + '~' + delayMaxSec + '秒' : '已关闭'}`);
        };

        // 恢复默认设置
        document.getElementById('azh-reset-settings').onclick = () => {
            CONFIG.concurrency = 2;
            CONFIG.randomDelay.enabled = true;
            CONFIG.randomDelay.min = 2000;
            CONFIG.randomDelay.max = 8000;

            document.getElementById('azh-setting-concurrency').value = 2;
            document.getElementById('azh-setting-random-delay').checked = true;
            document.getElementById('azh-setting-delay-min').value = 2;
            document.getElementById('azh-setting-delay-max').value = 8;
            document.getElementById('azh-sms-concurrency').value = 2;
            document.getElementById('azh-collection-concurrency').value = 2;
            const repayConc = document.getElementById('azh-repayment-concurrency');
            if (repayConc) repayConc.value = 2;

            Log.info('设置已恢复默认值');
            alert('已恢复默认设置：并发2，延迟2~8秒');
        };

        // 同步 CONFIG 默认值到设置面板和各模块输入框
        document.getElementById('azh-setting-concurrency').value = CONFIG.concurrency;
        document.getElementById('azh-setting-random-delay').checked = CONFIG.randomDelay.enabled;
        document.getElementById('azh-setting-delay-min').value = CONFIG.randomDelay.min / 1000;
        document.getElementById('azh-setting-delay-max').value = CONFIG.randomDelay.max / 1000;
        document.getElementById('azh-sms-concurrency').value = CONFIG.concurrency;
        document.getElementById('azh-collection-concurrency').value = CONFIG.concurrency;
        const repayConcEl = document.getElementById('azh-repayment-concurrency');
        if (repayConcEl) repayConcEl.value = CONFIG.concurrency;

        initialized = true;
        Log.success('初始化完成！爱租机小助手已就绪');
        return { success: true };
    }

    // 对外暴露初始化入口（加载器调用，密码在远程校验）
    window.AiZuJiHelperInit = function(password) {
        if (password !== CONFIG.password) {
            return { success: false, message: '密码错误，请重试' };
        }
        return init();
    };

    // 暴露到全局，方便调试
    window.AiZuJiHelper = {
        CONFIG,
        Log,
        Utils,
        UI,
        API,
        SmsModule,
        CollectionModule,
        RepaymentModule,
    };
})();
