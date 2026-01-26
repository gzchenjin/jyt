document.addEventListener('DOMContentLoaded', function() {
    // ======================= DATA INITIALIZATION =======================
    // 1. 定义全局变量，初始为空，等待数据加载
    let PM_DATA = [];
    let EMAIL_DATA = {};

    // 2. 异步加载 data.json 文件
    fetch('data.json')
        .then(response => {
            if (!response.ok) {
                throw new Error('网络响应异常: ' + response.statusText);
            }
            return response.json();
        })
        .then(data => {
            // 3. 将加载的数据赋值给全局变量
            PM_DATA = data.pm_data;
            EMAIL_DATA = data.email_data;
            console.log("人员与邮件数据加载成功！");
        })
        .catch(error => {
            console.error('数据加载失败:', error);
            alert("⚠️ 警告：人员基础数据加载失败，查询与邮件功能可能无法使用。\n请检查网络或 data.json 文件是否存在。");
        });

    // ======================= 2. 帮助文档加载 (help-content.html) =======================
    // 找到帮助文档的容器
    const helpContainer = document.querySelector('.help-content');

    if (helpContainer) {
        fetch('help-content.html')
            .then(response => {
                if (!response.ok) throw new Error(response.statusText);
                return response.text();
            })
            .then(html => {
                // 1. 注入 HTML
                helpContainer.innerHTML = html;
                
                // 2. 重新渲染数学公式 (MathJax)
                if (window.MathJax && window.MathJax.typesetPromise) {
                    MathJax.typesetPromise([helpContainer]).catch(err => console.log('MathJax渲染错误: ' + err.message));
                }

                // ============================================================
                // 【新增功能】修复目录点击导致整个页面滚动的问题
                // ============================================================
                // 选取帮助文档内所有以 # 开头的链接（即目录锚点）
                const tocLinks = helpContainer.querySelectorAll('a[href^="#"]');
                
                tocLinks.forEach(link => {
                    link.addEventListener('click', function(e) {
                        e.preventDefault(); // 1. 阻止浏览器默认的“整页跳转”行为
                        
                        // 获取目标元素的 ID (去掉 # 号)
                        const targetId = this.getAttribute('href').substring(1); 
                        const targetElement = document.getElementById(targetId);
                        
                        if (targetElement) {
                            // 2. 计算目标元素相对于 helpContainer 顶部的距离
                            // offsetTop 是元素相对于父定位元素的距离
                            const targetPosition = targetElement.offsetTop;
                            
                            // 3. 只滚动 helpContainer 容器内部，而不是整个 window
                            helpContainer.scrollTo({
                                top: targetPosition - 20, // 减去 20px 留点呼吸空间，避免顶格太紧
                                behavior: 'smooth'        // 平滑滚动效果
                            });
                        }
                    });
                });
                // ============================================================
            })
            .catch(err => {
                console.error("帮助文档加载失败:", err);
                helpContainer.innerHTML = "<p style='color:red; padding:20px;'>加载帮助文档失败，请检查 help-content.html 是否存在。</p>";
            });
    }

    // ======================= ELEMENT SELECTORS (DOM Cache) =======================
    const DOMElements = {
        // Main Form
        mainForm: document.getElementById('main-form'),
        ironTriangleInput: document.getElementById('ironTriangleInput'),
        projectLevel: document.getElementById('projectLevel'),
        budgetAmount: document.getElementById('budgetAmount'),
        
        // UI Interaction Elements
        procurement: document.getElementById('procurement'),
        procurementAmount: document.getElementById('procurementAmount'),
        procurementRiskRow: document.getElementById('procurementRiskRow'),
        coreCapability: document.getElementById('coreCapability'),
        coreCapabilityWarning: document.getElementById('coreCapabilityWarning'),
        capacityType: document.getElementById('capacityType'),
        deliveryDetailsTable: document.getElementById('deliveryDetailsTable'),
        
        SJ_projectCooperationNeeded: document.getElementById('SJ_projectCooperationNeeded'),
        SJ_projectCooperationAssessmentRow: document.getElementById('SJ_projectCooperationAssessmentRow'),
        SJ_preInvestmentNeeded: document.getElementById('SJ_preInvestmentNeeded'),
        SJ_preInvestmentDetailsRow: document.getElementById('SJ_preInvestmentDetailsRow'),
        
        TB_biddingMethod: document.getElementById('TB_biddingMethod'),
        TB_bidOpeningDate: document.getElementById('TB_bidOpeningDate'),
        TB_bidResponseDate: document.getElementById('TB_bidResponseDate'),
        TB_biddingRiskRow: document.getElementById('TB_biddingRiskRow'),
        TB_projectCooperationNeeded: document.getElementById('TB_projectCooperationNeeded'),
        TB_projectCooperationAssessmentRow: document.getElementById('TB_projectCooperationAssessmentRow'),
        TB_isPrimarySystem: document.getElementById('TB_isPrimarySystem'),
        TB_securityAssessmentRow: document.getElementById('TB_securityAssessmentRow'),
        
        JD_biddingMethod: document.getElementById('JD_biddingMethod'),
        JD_bidOpeningDate: document.getElementById('JD_bidOpeningDate'),
        JD_bidResponseDate: document.getElementById('JD_bidResponseDate'),
        JD_awardDate: document.getElementById('JD_awardDate'),
        
        // Help Panel
        helpContent: document.querySelector('.help-content'),

        // Modals
        reportModal: document.getElementById('report-modal'),
        reportTitle: document.getElementById('report-title'),
        reportText: document.getElementById('report-text'),

        pmQueryModal: document.getElementById('pm-query-modal'),
        pmAutoResult: document.getElementById('pm-auto-result'),
        pmSearchInput: document.getElementById('pm-search-input'),
        pmDeptFilter: document.getElementById('pm-dept-filter'),
        pmFullTableBody: document.getElementById('pm-full-table').querySelector('tbody'),

        attendeesModal: document.getElementById('attendees-modal'),
        attendeesForm: document.querySelector('#attendees-modal .attendees-form'),
        attendeesOutput: document.getElementById('attendees-output'),

        emailModal: document.getElementById('email-modal'),
        emailChecklist: document.getElementById('email-checklist'),
        emailOutput: document.getElementById('email-output'),
    };

    // ======================= INITIALIZATION =======================
    


    // Create Delivery Details Table Rows
    const tableBody = DOMElements.deliveryDetailsTable.querySelector('tbody');
    const verticalHeaders = ['牵头交付事业部', ...Array.from({length: 7}, (_, i) => `协助交付事业部${i + 1}`)];
    const deptOptions = ["", "IT系统事业部", "大数据AI应用事业部", "数字政府事业部/社会治理大数据研究院广州分院", "云网事业部", "智呼事业部", "智慧企业集成事业部/工业主研院", "智慧网络运营事业部", "智慧业财事业部"];

    verticalHeaders.forEach(headerText => {
        const row = tableBody.insertRow();
        row.innerHTML = `
            <th>${headerText}</th>
            <td><select class="table-dept-select">${deptOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}</select></td>
            <td><input type="text"></td>
            <td><input type="text"></td>
            <td><input type="text"></td>
        `;
    });

    // ======================= UI INTERACTIVITY LOGIC =======================
    
    function setupConditionalVisibility(selectId, rowId) {
        const select = document.getElementById(selectId);
        const row = document.getElementById(rowId);
        select.addEventListener('change', () => {
            const shouldShow = (select.value === '是');
            row.style.display = shouldShow ? '' : 'none';
            if (!shouldShow) {
                const textarea = row.querySelector('textarea');
                if (textarea) textarea.value = '';
            }
        });
    }

    DOMElements.procurement.addEventListener('change', () => {
        const isProcurement = DOMElements.procurement.value === '是';
        DOMElements.procurementAmount.disabled = !isProcurement;
        DOMElements.coreCapability.disabled = !isProcurement;
        DOMElements.procurementRiskRow.style.display = isProcurement ? '' : 'none';
        if (!isProcurement) {
            DOMElements.coreCapability.value = '';
            DOMElements.coreCapabilityWarning.style.display = 'none';
            DOMElements.procurementAmount.value = '';
            document.getElementById('procurementRisk').value = '';
        }
    });

    DOMElements.coreCapability.addEventListener('change', () => {
        DOMElements.coreCapabilityWarning.style.display = (DOMElements.coreCapability.value === '是') ? '' : 'none';
    });

    DOMElements.capacityType.addEventListener('change', () => {
        const showTable = ['多产能', '大集成'].includes(DOMElements.capacityType.value);
        DOMElements.deliveryDetailsTable.style.display = showTable ? 'table' : 'none';
    });

    setupConditionalVisibility('SJ_projectCooperationNeeded', 'SJ_projectCooperationAssessmentRow');
    setupConditionalVisibility('SJ_preInvestmentNeeded', 'SJ_preInvestmentDetailsRow');
    setupConditionalVisibility('TB_projectCooperationNeeded', 'TB_projectCooperationAssessmentRow');
    setupConditionalVisibility('TB_isPrimarySystem', 'TB_securityAssessmentRow');

    function updateBiddingFields(methodSelectId) {
        const method = document.getElementById(methodSelectId).value;
        const isTB = methodSelectId.startsWith('TB');
        const openDateEl = document.getElementById(isTB ? 'TB_bidOpeningDate' : 'JD_bidOpeningDate');
        const responseDateEl = document.getElementById(isTB ? 'TB_bidResponseDate' : 'JD_bidResponseDate');
        const awardDateEl = document.getElementById('JD_awardDate');
        const riskRowEl = document.getElementById('TB_biddingRiskRow');

        const openMethods = ["公开招标", "邀请招标", "比选"];
        const responseMethods = ["单一来源", "询价", "竞争性谈判"];
        
        const isOpen = openMethods.includes(method);
        const isResponse = responseMethods.includes(method);
        
        openDateEl.disabled = !isOpen;
        responseDateEl.disabled = !isResponse;
        if(isTB && riskRowEl) riskRowEl.style.display = isOpen ? '' : 'none';
        if(!isTB && awardDateEl) awardDateEl.disabled = !isOpen;
    };
    
    DOMElements.TB_biddingMethod.addEventListener('change', () => updateBiddingFields('TB_biddingMethod'));
    DOMElements.JD_biddingMethod.addEventListener('change', () => updateBiddingFields('JD_biddingMethod'));

    // ======================= FIELD SYNCHRONIZATION =======================
    const syncPairs = {
        'SJ_grossMargin': 'TB_grossMargin', 'TB_grossMargin': 'SJ_grossMargin',
        'SJ_projectCooperationAssessment': 'TB_projectCooperationAssessment', 'TB_projectCooperationAssessment': 'SJ_projectCooperationAssessment',
        'SJ_projectCooperationNeeded': 'TB_projectCooperationNeeded', 'TB_projectCooperationNeeded': 'SJ_projectCooperationNeeded',
        'TB_businessType': 'JD_businessType', 'JD_businessType': 'TB_businessType',
        'TB_biddingMethod': 'JD_biddingMethod', 'JD_biddingMethod': 'TB_biddingMethod',
        'JD_deliveryPeriod': 'TB_deliveryPeriod', 'TB_deliveryPeriod': 'JD_deliveryPeriod',
        'JD_deliveryRisk': 'TB_deliveryRisk', 'TB_deliveryRisk': 'JD_deliveryRisk',
        'JD_maintenanceRequirements': 'TB_maintenanceRequirements', 'TB_maintenanceRequirements': 'JD_maintenanceRequirements',
        'JD_trialRun': 'TB_trialRun', 'TB_trialRun': 'JD_trialRun',
        'JD_maintenanceAssessment': 'TB_maintenanceAssessment', 'TB_maintenanceAssessment': 'JD_maintenanceAssessment',
        'JD_testingRequirements': 'TB_testingRequirements', 'TB_testingRequirements': 'JD_testingRequirements',
    };
    
    let isSyncing = false;
    
    Object.keys(syncPairs).forEach(sourceId => {
        const sourceEl = document.getElementById(sourceId);
        sourceEl.addEventListener('change', () => { // Use 'change' for selects, works for inputs too
            if (isSyncing) return;
            isSyncing = true;
            const targetEl = document.getElementById(syncPairs[sourceId]);
            if (targetEl) {
                targetEl.value = sourceEl.value;
                targetEl.dispatchEvent(new Event('change', { bubbles: true }));
            }
            isSyncing = false;
        });
        sourceEl.addEventListener('input', () => { // Use 'input' for live typing in text fields
            if (isSyncing || sourceEl.tagName === 'SELECT') return;
             isSyncing = true;
            const targetEl = document.getElementById(syncPairs[sourceId]);
            if (targetEl) targetEl.value = sourceEl.value;
            isSyncing = false;
        });
    });

    // ======================= CORE LOGIC (from logic_1.py) =======================

    function _safeFloat(value, defaultValue = 0.0) {
        const num = parseFloat(value);
        return isNaN(num) ? defaultValue : num;
    }

    function _ensurePeriod(text) {
        text = String(text || '').trim();
        if (!text) return "";
        return (text.endsWith('。') || text.endsWith('.')) ? text : text + '。';
    }

    function _formatDate(dateStr) {
        try {
            const dt = new Date(dateStr);
            if (isNaN(dt.getTime())) return dateStr;
            return `${dt.getFullYear()}年${dt.getMonth() + 1}月${dt.getDate()}日`;
        } catch (e) {
            return dateStr;
        }
    }



    function _formatProjectRoles(inputText) {
        if (!inputText) return "【铁三角信息未填写】。";
        const roles = {};
        // --- 修改后的正则表达式 ---
        const pattern = /(.+?)\s*[:：]\s*(.+?)\s*[(（](.+?)[)）]/gs; 
        // --- 修改结束 ---
        let match;
        while ((match = pattern.exec(inputText.trim())) !== null) {
            // Trim potential extra spaces around names and departments
            roles[match[1].trim()] = { name: match[2].trim(), department: match[3].trim() };
        }
        const getInfo = (role) => roles[role] || { name: "【待定】", department: "【待定】" };
        const pm = getInfo("项目经理"), sales = getInfo("销售经理"), solution = getInfo("方案经理"), delivery = getInfo("交付经理");
        // Ensure department is captured correctly even if name extraction failed slightly
        if (!pm.department && roles["项目经理"]?.department) pm.department = roles["项目经理"].department;
        if (!sales.department && roles["销售经理"]?.department) sales.department = roles["销售经理"].department;
        if (!solution.department && roles["方案经理"]?.department) solution.department = roles["方案经理"].department;
        if (!delivery.department && roles["交付经理"]?.department) delivery.department = roles["交付经理"].department;
        
        return `项目经理是${pm.name}(${pm.department})，销售经理是${sales.name}(${sales.department})，方案经理是${solution.name}(${solution.department})，交付经理是${delivery.name}(${delivery.department})。`;
    }

    function gatherFormData() {
        const data = {};
        const form = DOMElements.mainForm;
        const inputs = form.querySelectorAll('input, select, textarea');
        inputs.forEach(input => {
            if (input.id) data[input.id] = input.value;
        });
        
        const table_data = [];
        const rows = DOMElements.deliveryDetailsTable.querySelector('tbody').rows;
        for(let i=0; i<rows.length; i++) {
            const row = rows[i];
            const dept = row.cells[1].querySelector('select').value;
            if(!dept) continue;

            table_data.push({
                '类型': row.cells[0].innerText,
                '事业部': dept,
                '项目经理': row.cells[2].querySelector('input').value,
                '交付内容': row.cells[3].querySelector('input').value,
                '预算（万元）': row.cells[4].querySelector('input').value
            });
        }
        data['deliveryDetails'] = table_data;
        return data;
    }

    function _calculateCommonLogic(formData) {
        const budgetYuan = _safeFloat(formData.budgetAmount);
        const procurementYuan = _safeFloat(formData.procurementAmount);
        const budgetWan = budgetYuan / 10000;
        const procurementWan = procurementYuan / 10000;
        const procurementRatio = (budgetWan !== 0) ? (procurementWan / budgetWan) : 0.0;
        
        let leadDepartment = "【未知事业部】";
        const tjsRawText = formData.ironTriangleInput || '';
        // --- 修改后的正则表达式 ---
        const pmPattern = /项目经理\s*[:：]\s*(?:.+?)\s*[(（](.+?)[)）]/s;
        // --- 修改结束 ---
        const pmMatch = tjsRawText.match(pmPattern);
        if (pmMatch && pmMatch[1]) {
             leadDepartment = pmMatch[1].trim();
        }

        const tjsSummary = _formatProjectRoles(tjsRawText); // Uses the updated function from Step 1

        // Delivery details summaries (logic remains the same)

        let assistDepartmentsSummary = '';
        let deliverySummary = '';
        const deliveryDetails = formData.deliveryDetails || [];
        if (formData.capacityType !== '单产能' && deliveryDetails.length > 0) {
            const assistDepts = deliveryDetails
                .filter(row => row['类型'] !== '牵头交付事业部' && row['事业部'])
                .map(row => row['事业部'].trim());
            if (assistDepts.length > 0) assistDepartmentsSummary = `${assistDepts.join('、')}协助交付。`;

            deliverySummary = '\n' + deliveryDetails.map((row, index) => {
                let currentPart = `${row['事业部']}`;
                if (row['交付内容']) currentPart += `负责交付${row['交付内容']}`;
                const budgetVal = _safeFloat(row['预算（万元）']);
                if (budgetVal > 0) currentPart += `，预算${budgetVal.toFixed(2)}万元`;
                if (row['项目经理']) {
                    const managerRole = row['类型'] === '牵头交付事业部' ? '项目经理' : '子项目经理';
                    currentPart += `，${managerRole}是${row['项目经理']}`;
                }
                return `${index + 1}）${currentPart}`;
            }).join('；\n') + '。';
        }

        return {
            budget_in_wan: budgetWan,
            procurement_in_wan: procurementWan,
            procurement_ratio: procurementRatio,
            lead_department: leadDepartment,
            tjs_summary: tjsSummary,
            assist_departments_summary: assistDepartmentsSummary,
            delivery_summary: deliverySummary
        };
    }
    
    // --- Generate Opportunity Minutes ---
    function generateOpportunityMinutes(formData) {
        const common = _calculateCommonLogic(formData);
        const temp = {...formData, ...common};
        const getVal = (key, placeholder) => temp[key]?.trim() || placeholder;

        const projectName = getVal('projectName', "【请补充项目名称】");
        const businessCode = getVal('businessCode', "【请补充项目商机编码】");
        const constructionContent = getVal('constructionContent', "【请补充建设内容】");
        const capacityType = getVal('capacityType', "【请选择产能能力】");
        const projectLevel = getVal('projectLevel', "【请选择项目级别】");
        const SJ_projectRisk = `本项目存在风险：\n${getVal('SJ_projectRisk', "【请补充项目风险（商机）】")}`;
        const OT_risk = getVal('OT_risk');
        
        const clientDescription = !temp.contractClient  ? "【请补充签约客户】"
          : `${temp.contractClient}` 
          + ((!temp.endClient || temp.endClient === temp.contractClient) ? '。'
          : `，最终客户是${temp.endClient}。`);

        let procurement_text = temp.procurement === '是' ? `涉及外采，外采预算${temp.procurement_in_wan.toFixed(2)}万元（含税），`
            : temp.procurement === '否' ? '不涉及外采，' : '【请选择是否后向外采】，';
        
        let procurementRisk = temp.procurement === '是' ? getVal('procurementRisk', "【请补充外采风险】")
            : temp.procurement === '否' ? "本项目不涉及外采" : "【请评估是否涉及外采】";
        
        let SJ_projectCooperationAssessment = temp.SJ_projectCooperationNeeded === '是' ? getVal('SJ_projectCooperationAssessment', "【请补充项目合作评估】")
            : temp.SJ_projectCooperationNeeded === '否' ? "本项目不涉及项目合作" : "【请评估是否项目合作】";
        
        let SJ_preInvestmentDetails = temp.SJ_preInvestmentNeeded === '是' ? getVal('SJ_preInvestmentDetails', "【请补充预投入情况】")
            : temp.SJ_preInvestmentNeeded === '否' ? "本项目不涉及预投入情况" : "【请评估是否涉及预投入】";

        const SJ_atomicCapability = getVal('SJ_atomicCapability', "【请补充原子能力评估】");



        
        const key1_text = `本项目是${projectName}，项目签约客户是${clientDescription}项目建设内容为${constructionContent}`;
        const key2_text = `本项目属于${capacityType}${projectLevel}项目，由${temp.lead_department}牵头，${temp.assist_departments_summary}${temp.tjs_summary}${temp.delivery_summary}`;
        const key3_text = `本项目预算${temp.budget_in_wan.toFixed(2)}万元（含税），${procurement_text}毛利率预估${_safeFloat(temp.SJ_grossMargin).toFixed(2)}%（不含税），利润率预估${_safeFloat(temp.SJ_netMargin).toFixed(2)}%（不含税）`;

        let output = `${projectName}（${businessCode}）\n\n商机评估会\n`;
        const points = [key1_text, key2_text, key3_text, SJ_projectRisk, procurementRisk, SJ_projectCooperationAssessment, SJ_preInvestmentDetails, SJ_atomicCapability, OT_risk];
        
        let counter = 1;
        points.forEach(p => {
            if (p && String(p).trim()) {
                output += `${counter}、${_ensurePeriod(p)}\n`;
                counter++;
            }
        });
        output += "综合评估各要素，铁三角成员评估跟进此商机。";
        return output;
    }

    // --- Generate Bidding Minutes ---
    function generateBiddingMinutes(formData) {
        const common = _calculateCommonLogic(formData);
        const temp = {...formData, ...common};
        const getVal = (key, placeholder) => temp[key]?.trim() || placeholder;

        const projectName = getVal('projectName', "【请补充项目名称】");
        const businessCode = getVal('businessCode', "【请补充项目商机编码】");
        const TB_businessType = getVal('TB_businessType', "【请补充业务类型】");
        const constructionContent = getVal('constructionContent', "【请补充建设内容】");
        const capacityType = getVal('capacityType', "【请选择产能能力】");
        const projectLevel = getVal('projectLevel', "【请选择项目级别】");
        const TB_deliveryPeriod = getVal('TB_deliveryPeriod', "【请补充交付周期】");
        const TB_deliveryRisk = getVal('TB_deliveryRisk', "【请补充交付风险】");
        const TB_maintenanceRequirements = getVal('TB_maintenanceRequirements', "【请补充运维要求】");
        const TB_maintenanceAssessment = getVal('TB_maintenanceAssessment', "【请补充运维服务评估意见】");
        const TB_financialAssessment = getVal('TB_financialAssessment', "【请补充财务评估】");
        const TB_testingRequirements = getVal('TB_testingRequirements', "【请补充等保测评、第三方测评要求】");
        const TB_trialRun = getVal('TB_trialRun', "【请补充试运行情况】");
        const OT_risk = getVal('OT_risk');
        const TB_biddingRisk = getVal('TB_biddingRisk', "【请补充招投标风险评估】");

        const clientDescription = !temp.contractClient  ? "【请补充签约客户】"
          : `${temp.contractClient}` 
          + ((!temp.endClient || temp.endClient === temp.contractClient) ? '。'
          : `，最终客户是${temp.endClient}。`);
        

        const bidMethodsWithRisk = ["公开招标", "邀请招标", "比选"];
        const bid_method_desc = (method => {
            if (!method) return "【请选择投标方式】";
            const entity = getVal('TB_biddingEntity', '【请选择投标主体】');
            if (bidMethodsWithRisk.includes(method)) return `${method}方式，拟以${entity}名义投标，开标时间为${_formatDate(temp.TB_bidOpeningDate)}`;
            if (["单一来源", "询价", "竞争性谈判"].includes(method)) return `${method}方式，拟以${entity}名义应答，应答时间为${_formatDate(temp.TB_bidResponseDate)}`;
            return method;
        })(temp.TB_biddingMethod);

        const risk_assessment_desc = bidMethodsWithRisk.includes(temp.TB_biddingMethod) ? TB_biddingRisk
            : (temp.TB_biddingMethod ? `本项目客户采用${temp.TB_biddingMethod}方式，不涉及招投标风险` : "【请选择投标方式】");

        const procurement_text = temp.procurement === '是' ? `涉及外采，外采预算${temp.procurement_in_wan.toFixed(2)}万元（含税），外采占比${temp.procurement_ratio.toLocaleString(undefined, {style: 'percent', minimumFractionDigits: 2})}，`
            : temp.procurement === '否' ? '不涉及外采，' : '【请选择是否后向外采】，';

        const procurementRisk = temp.procurement === '是' ? getVal('procurementRisk', "【请补充外采风险】")
            : temp.procurement === '否' ? "本项目不涉及外采" : "【请评估是否涉及外采】";

        const TB_projectCooperationAssessment = temp.TB_projectCooperationNeeded === '是' ? getVal('TB_projectCooperationAssessment', "【请补充项目合作评估】")
            : temp.TB_projectCooperationNeeded === '否' ? "本项目不涉及项目合作" : "【请评估是否项目合作】";
            
        const TB_securityAssessment = temp.TB_isPrimarySystem === '是' ? getVal('TB_securityAssessment', "【请补充网络和信息安全评估】")
            : temp.TB_isPrimarySystem === '否' ? "本项目不涉及亿迅主责系统" : "【请评估是否亿迅主责系统】";



        // 先处理 constructionContent，移除末尾可能自带的句号
        const trimmedContent = constructionContent.endsWith('。') 
            ? constructionContent.slice(0, -1) 
            : constructionContent;


        const key1_text = `本项目为${TB_businessType}项目，项目签约客户是${clientDescription}客户计划采用${bid_method_desc}。`;
        const key2_text = `项目建设内容为：${trimmedContent}。由${temp.lead_department}牵头，${temp.assist_departments_summary}${temp.tjs_summary}${temp.delivery_summary}`;
        const key3_text = `本项目预算${temp.budget_in_wan.toFixed(2)}万元（含税），属于${capacityType}${projectLevel}项目，${procurement_text}毛利率预估${_safeFloat(temp.TB_grossMargin).toFixed(2)}%（不含税）。`;
        const key4_text = `本项目交付要求：交付周期为${TB_deliveryPeriod}，${TB_deliveryRisk}`;

        let output = `${projectName}(${businessCode})\n\n投标评估会\n一、项目基本信息\n1、${key1_text}\n2、${key2_text}\n3、${key3_text}\n二、风险及应对措施\n`;
        
        const risk_points = [risk_assessment_desc, key4_text, procurementRisk, TB_projectCooperationAssessment, TB_maintenanceRequirements, TB_securityAssessment, TB_maintenanceAssessment, TB_financialAssessment, TB_testingRequirements, TB_trialRun, OT_risk];
        let counter = 1;
        risk_points.forEach(p => {
            if (p && String(p).trim()) {
                output += `${counter}、${_ensurePeriod(p)}\n`;
                counter++;
            }
        });
        
        output += "三、会议结论\n综合评估各要素，铁三角评估可参与此项目投标，并根据内控审批权限作投标审批决策。";
        return output;
    }

    // --- Generate Kickoff Minutes ---
    function generateKickoffMinutes(formData) {
        const common = _calculateCommonLogic(formData);
        const temp = {...formData, ...common};
        const getVal = (key, placeholder) => temp[key]?.trim() || placeholder;

        const projectName = getVal('projectName', "【请补充项目名称】");
        const businessCode = getVal('businessCode', "【请补充项目商机编码】");
        const JD_businessType = getVal('JD_businessType', "【请补充业务类型】");
        const constructionContent = getVal('constructionContent', "【请补充建设内容】");
        const capacityType = getVal('capacityType', "【请选择产能能力】");
        const projectLevel = getVal('projectLevel', "【请选择项目级别】");
        const JD_deliveryPeriod = getVal('JD_deliveryPeriod', "【请补充交付周期】");
        const JD_deliveryRisk = getVal('JD_deliveryRisk', "【请补充交付风险】");
        const JD_maintenanceRequirements = getVal('JD_maintenanceRequirements', "【请补充运维要求】");
        const JD_trialRun = getVal('JD_trialRun', "【请补充试运行情况】");
        const JD_maintenanceAssessment = getVal('JD_maintenanceAssessment', "【请补充运维服务评估意见】");
        const JD_testingRequirements = getVal('JD_testingRequirements', "【请补充等保测评、第三方测评要求】");
        const OT_risk = getVal('OT_risk');

        const clientDescription = !temp.contractClient  ? "【请补充签约客户】"
          : `${temp.contractClient}` 
          + ((!temp.endClient || temp.endClient === temp.contractClient) ? '。'
          : `，最终客户是${temp.endClient}。`);

        // ========================[ 这里是您确认的最新逻辑 ]========================
        const bid_method_desc = (method => {
            if (!method) return "【请选择投标方式】";
            
            const open_date = _formatDate(temp.JD_bidOpeningDate);
            const award_date = _formatDate(temp.JD_awardDate);
            const signing_date = _formatDate(temp.JD_signingDate);
            const response_date = _formatDate(temp.JD_bidResponseDate);
    
            if (["公开招标", "邀请招标", "比选"].includes(method)) {
                return `${method}方式，开标时间为${open_date}，中标时间为${award_date}，计划${signing_date}前完成签约`;
            }
            if (method === "单一来源") {
                return `${method}方式，无需招投标，${response_date}已完成应答，计划${signing_date}前完成签约`;
            }
            if (["原子能力下单", "订单方式"].includes(method)) {
                return `${method}方式，预计客户在${signing_date}前完成下单`;
            }
            if (["询价", "竞争性谈判"].includes(method)) {
                return `${method}方式，${response_date}已完成应答，预计客户在${signing_date}前完成签约`;
            }
            if (["电商采购", "直接采购"].includes(method)) {
                return `${method}方式，预计客户在${signing_date}前完成签约`;
            }
            return method; // 其他情况直接返回方法名
        })(temp.JD_biddingMethod);
        // ======================================================================

        const procurement_text = temp.procurement === '是' ? `涉及外采，外采预算${temp.procurement_in_wan.toFixed(2)}万元（含税），外采占比${temp.procurement_ratio.toLocaleString(undefined, {style: 'percent', minimumFractionDigits: 2})}。`
            : temp.procurement === '否' ? '不涉及外采。' : '【请选择是否后向外采】，';
            
        const procurementRisk = temp.procurement === '是' ? getVal('procurementRisk', "【请补充外采风险】")
            : temp.procurement === '否' ? "本项目不涉及外采" : "【请评估是否涉及外采】";

        const trimmedContent = constructionContent.endsWith('。') 
            ? constructionContent.slice(0, -1) 
            : constructionContent;
        
        const key1_text = `本项目为${JD_businessType}项目，项目签约客户是${clientDescription}客户采用${bid_method_desc}。`;
        const key2_text = `本项目预算${temp.budget_in_wan.toFixed(2)}万元（含税），${procurement_text}`;
        const key3_text = `项目建设内容为：${trimmedContent}。本项目属于${capacityType}${projectLevel}项目，由${temp.lead_department}牵头，${temp.assist_departments_summary}${temp.tjs_summary}${temp.delivery_summary}`;
        const key4_text = "1、项目售前资料交底：销售经理、方案经理已对项目所有售前的会议纪要、客户沟通记录、客户需求及交付要求等资料交接给交付经理、项目经理；\n2、项目投标资料交底：销售经理、方案经理已对招标文件、投标文件、技术规范书等资料交接给交付经理、项目经理；\n3、项目实施计划交底：项目经理已完成项目里程碑计划，各关键节点已有明确的交付成果要求，铁三角已确认该时间节点可行；\n4、项目干系人交底：销售经理已上传项目干系人清单，清单已包含客户（签约客户/最终客户）以及合作伙伴干系人的名单和联系方式，铁三角对项目干系人已知晓。";
        const key5_text = `本项目交付要求：交付周期为${JD_deliveryPeriod}，${JD_deliveryRisk}`;

        let output = `${projectName}（${businessCode}）\n\n项目交底会\n一、项目基本信息\n1、${key1_text}\n2、${key2_text}\n3、${key3_text}\n二、项目文件交底\n${key4_text}\n三、风险及应对举措\n`;
        
        const risk_points = [key5_text, JD_maintenanceRequirements, procurementRisk, JD_trialRun, JD_maintenanceAssessment, JD_testingRequirements, OT_risk];
        let counter = 1;
        risk_points.forEach(p => {
            if (p && String(p).trim()) {
                output += `${counter}、${_ensurePeriod(p)}\n`;
                counter++;
            }
        });

        output += "四、其他参会部门意见\n公共架构评估师意见详见会议纪要中【公共架构结论】部分。\n五、会议结论\n项目铁三角对项目情况、项目角色分工、项目计划及里程碑节点、项目风险及问题解决方案等内容均已了解清晰，交底完成，请项目组尽快完成合同签约。";
        return output;
    }
    
    // ======================= EVENT LISTENERS FOR BUTTONS =======================
    
    // 1. 商机评估
    document.getElementById('btn-gen-opportunity').addEventListener('click', () => {
        // --- 生成报告部分（保持原样，用原始数据）---
        const formData = gatherFormData(); 
        const report = generateOpportunityMinutes(formData);
        showReportDialog("商机评估会纪要", report);

        // --- 发送给服务器部分（修改这里）---
        //const data = gatherFormData();
        // 【核心修改】：直接修改项目名称，加上标记
        // 假设您的表单里项目名称字段叫 project_name
        //data.projectName = data.projectName + "【商机】"; 
        //saveDataToBackend(data); 
    });

    // 2. 投标评估
    document.getElementById('btn-gen-bidding').addEventListener('click', () => {
        const formData = gatherFormData();
        const report = generateBiddingMinutes(formData);
        showReportDialog("投标评估会纪要", report);

        //const data = gatherFormData();
        //data.projectName = data.projectName + "【投标】"; 
        //saveDataToBackend(data); 
    });
    
    // 3. 项目交底
    document.getElementById('btn-gen-kickoff').addEventListener('click', () => {
        const formData = gatherFormData();
        const report = generateKickoffMinutes(formData);
        showReportDialog("项目交底会纪要", report);

        //const data = gatherFormData();     
        //data.projectName = data.projectName + "【交底】"; 
        //saveDataToBackend(data); 
    });

    // ======================= MODAL/DIALOG HANDLING =======================
    
    function showReportDialog(title, text) {
        DOMElements.reportTitle.innerText = title;
        DOMElements.reportText.value = text;
        DOMElements.reportModal.style.display = 'flex';
    }
    document.getElementById('close-report-modal').addEventListener('click', () => DOMElements.reportModal.style.display = 'none');

    // --- PM Query Modal Logic ---
    document.getElementById('btn-query-pm').addEventListener('click', () => {
        DOMElements.pmFullTableBody.innerHTML = '';
        PM_DATA.forEach(pm => {
            const row = DOMElements.pmFullTableBody.insertRow();
            row.innerHTML = `<td>${pm.项目经理}</td><td>${pm.级别}</td><td>${pm.在职部门}</td>`;
        });
        const allDepts = [...new Set(PM_DATA.map(p => p.在职部门))].sort();
        DOMElements.pmDeptFilter.innerHTML = `<option>所有部门</option>` + allDepts.map(d => `<option>${d}</option>`).join('');
        
        performPMAutoSearch();
        filterPMTable();
        DOMElements.pmQueryModal.style.display = 'flex';
    });
    document.getElementById('close-pm-modal').addEventListener('click', () => DOMElements.pmQueryModal.style.display = 'none');
    
    function performPMAutoSearch() {
        const ironTriangleText = DOMElements.ironTriangleInput.value;
        // --- 修改后的正则表达式 ---
        const pattern = /项目经理\s*[:：]\s*(.+?)\s*[(（](.+?)[)）]/s;
        // --- 修改结束 ---
        const match = ironTriangleText.match(pattern);
        
        if (!match) {
            DOMElements.pmAutoResult.innerHTML = "未在铁三角信息中找到有效的项目经理信息。";
            return;
        }
        const pmName = match[1].trim();
        const pmDept = match[2].trim();
        const matches = PM_DATA.filter(p => p.项目经理 === pmName);
        
        if (matches.length === 0) {
            DOMElements.pmAutoResult.innerHTML = `匹配不成功：未在列表中找到名为 '${pmName}' 的项目经理。`;
            return;
        }
        matches.sort((a, b) => (a.在职部门 === pmDept ? -1 : 1) - (b.在职部门 === pmDept ? -1 : 1));
        let result_text = `<b>为 '${pmName}' 找到 ${matches.length} 个结果（部门匹配优先）：</b><br>`;
        result_text += matches.map(p => `&nbsp;&nbsp;- ${p.项目经理} (<b>${p.级别}</b>) - ${p.在职部门}`).join('<br>');
        DOMElements.pmAutoResult.innerHTML = result_text;
    }

    
    
    function filterPMTable() {
        const searchText = DOMElements.pmSearchInput.value.toLowerCase();
        const selectedDept = DOMElements.pmDeptFilter.value;
        const rows = DOMElements.pmFullTableBody.rows;
        for (let i = 0; i < rows.length; i++) {
            const name = rows[i].cells[0].textContent.toLowerCase();
            const dept = rows[i].cells[2].textContent;
            const nameMatch = !searchText || name.includes(searchText);
            const deptMatch = selectedDept === '所有部门' || dept === selectedDept;
            rows[i].style.display = (nameMatch && deptMatch) ? '' : 'none';
        }
    }
    DOMElements.pmSearchInput.addEventListener('input', filterPMTable);
    DOMElements.pmDeptFilter.addEventListener('change', filterPMTable);

    
    // --- Attendees Modal Logic ---
    document.getElementById('btn-show-attendees').addEventListener('click', () => {
        const attendeesForm = DOMElements.attendeesForm;
        attendeesForm.querySelector('#attendees-projectLevel').value = DOMElements.projectLevel.value;
        try {
            const budgetWan = parseFloat(DOMElements.budgetAmount.value) / 10000;
            attendeesForm.querySelector('#attendees-budgetAmount').value = isNaN(budgetWan) ? '' : budgetWan;
        } catch { attendeesForm.querySelector('#attendees-budgetAmount').value = ''; }
        attendeesForm.querySelector('#attendees-procurement').value = DOMElements.procurement.value;
        attendeesForm.querySelector('#attendees-cooperation').value = DOMElements.SJ_projectCooperationNeeded.value || DOMElements.TB_projectCooperationNeeded.value;
        attendeesForm.querySelector('#attendees-primarySystem').value = DOMElements.TB_isPrimarySystem.value;
        
        updateAttendeesList();
        DOMElements.attendeesModal.style.display = 'flex';
    });
    document.getElementById('close-attendees-modal').addEventListener('click', () => DOMElements.attendeesModal.style.display = 'none');
    
    function updateAttendeesList() {
        const form = DOMElements.attendeesForm;
        const getV = (id) => form.querySelector(`#${id}`).value;
        
        let roles = {};
        // --- 修改后的正则表达式 ---
        const pattern = /(.+?)\s*[:：]\s*(.+?)\s*[(（](.+?)[)）]/gs;
        // --- 修改结束 ---
        const ironTriangleText = DOMElements.ironTriangleInput.value.trim();
        let match;
        while ((match = pattern.exec(ironTriangleText)) !== null) {
             roles[match[1].trim()] = { name: match[2].trim(), department: match[3].trim() };
        }
        
        const getRoleString = roleName => {
            const info = roles[roleName];
            return info && info.name && info.department ? `${roleName}：${info.name}【${info.department}】` : roleName;
        };
        
        const all = {
            '项目经理': getRoleString("项目经理"), 
            '销售经理': getRoleString("销售经理"), 
            '方案经理': getRoleString("方案经理"), 
            '交付经理': getRoleString("交付经理"),
            '法律风险评估': "法律风险评估: 李屹【法律合规部】", 
            '项目合作': "项目合作评估: 黄曦楠/许仲华【市场及渠道支撑部（标前）】",
            '采购评估': "采购评估: 梁其容/罗晓纯【采购部】", 
            '网信安评估': "网信安评估: 吴中华/陆艺阳【运营管理部/研发与质量管理中心、智慧网络运营事业部】",
            '运维服务评估意见': "运维服务评估: 熊俊伟, 蒋朝豪【运营管理部/研发与质量管理中心】",
            '公共架构评估': (() => {
                const archMap = {
                    "IT系统事业部": "王沛文、高航",
                    "大数据AI应用事业部": "许智洋",
                    "数字政府事业部/社会治理大数据研究院广州分院": "李佳鑫、许伟明",
                    "云网事业部": "许智洋",
                    "智呼事业部": "郑辉",
                    "智慧企业集成事业部/工业主研院": "郑辉、陈家辉",
                    "智慧网络运营事业部": "周宏江、高航、许伟明",
                    "智慧业财事业部": "王沛文",
                };
                const pmDept = roles["项目经理"]?.department;
                return `公共架构评估师： ${archMap[pmDept] || ''}【运营管理部/研发与质量管理中心】`;
            })()
        };

        const budget = _safeFloat(getV('attendees-budgetAmount'));
        const procurement = getV('attendees-procurement'); // 获取是否涉及外采
        const level = getV('attendees-projectLevel');       // 获取项目级别
        const industry = getV('attendees-industryType');     // [新增] 行业/电信
        const cooperation = getV('attendees-cooperation'); // [关键] 项目合作/变更
        

        if (budget >= 50) {
            // 首先判断：如果不涉及外采 且 项目级别是 B类 或 C类
            if (procurement === '否' && (level === 'B类' || level === 'C类')) {
                all['财务评估'] = "财务评估: 刘椰韵【财务部】";
            } else {
                // 其他所有预算大于等于50万的情况（涉及外采，或者项目是A类）
                all['财务评估'] = "财务评估: 戴亮【财务部】";
            }
        }
        // 注意：如果预算小于50万，则 all['财务评估'] 不会被赋值，即不邀请财务参会。
        
        const required = {
            "商机评估会": ["项目经理", "销售经理", "方案经理", "交付经理"],
            "投标评估会": ["项目经理", "销售经理", "方案经理", "交付经理", "运维服务评估意见", "财务评估"],
            "项目交底会": ["项目经理", "销售经理", "方案经理", "交付经理", "公共架构评估"],
            "商机、投标评估会": ["项目经理", "销售经理", "方案经理", "交付经理", "运维服务评估意见", "财务评估"],
        };

        const optionalFlags = {
            "项目合作": getV('attendees-cooperation') === '是',
            "采购评估": getV('attendees-procurement') === '是',
            "网信安评估": getV('attendees-primarySystem') === '是',
            "法律风险评估": getV('attendees-legalRisk') === '是',
        };

        const meetingType = getV('attendees-meetingType');
        let final = (required[meetingType] || []).map(role => all[role]).filter(Boolean);
        Object.keys(optionalFlags).forEach(role => {
            if (optionalFlags[role] && all[role]) final.push(all[role]);
        });
        
        // --- [修改] 4. 领导参会提示 ---
        // 条件：
        // 1. 行业
        // 2. A类
        // 3. 涉及投标会
        // 4. (外采是“是” OR 项目合作是“是”)  <-- 这里用了 || (或)
        if (industry === '行业' && level === 'A类' && meetingType.includes('投标') && 
           (procurement === '是' || cooperation === '是')) {
            final.push("⚠️ 行业域需外采A类项目：邀请渠道部门领导、涉及外采的产能部门领导以及市场及渠道支撑部领导（正职或分管该业务的副职）共同参会");
        }



        DOMElements.attendeesOutput.value = [...new Set(final)].join('\n');
    }
    DOMElements.attendeesForm.addEventListener('change', updateAttendeesList);
    DOMElements.attendeesForm.addEventListener('input', updateAttendeesList);

    // --- Email Modal Logic ---
    document.getElementById('btn-email').addEventListener('click', () => {
        repopulateEmailDepartments();
        DOMElements.emailModal.style.display = 'flex';
    });
    document.getElementById('close-email-modal').addEventListener('click', () => DOMElements.emailModal.style.display = 'none');
    
    const repopulateEmailDepartments = () => {
        DOMElements.emailChecklist.innerHTML = '';
        let depts = new Set();
        
        // --- 1. 获取铁三角文本 ---
        const tjsText = DOMElements.ironTriangleInput.value;
        console.log("正在分析铁三角文本...");

        // ============================================================
        // [核心修复] 超强容错正则
        // 解释：
        // 1. 销售经理     -> 锚点
        // 2. [:：]       -> 兼容 中文冒号 或 英文冒号
        // 3. \s* -> 兼容 任意数量的换行、空格、制表符
        // 4. [^(\（]*?    -> 非贪婪匹配，跳过人名 (直到遇到左括号为止)
        // 5. [(\（]      -> 兼容 中文左括号（ 或 英文左括号 (
        // 6. (.+?)       -> 【捕获目标】部门名称
        // 7. [)\）]      -> 兼容 中文右括号） 或 英文右括号 )
        // ============================================================
        const salesRegex = /销售经理[:：]\s*[^(\（]*?[(\（](.+?)[)\）]/;
        const salesMatch = tjsText.match(salesRegex);

        if (salesMatch) {
            const deptName = salesMatch[1].trim();
            console.log("✅ 成功抓取销售部门:", deptName);
            
            // 只要抓到了就添加，不管 EMAIL_DATA 里有没有配置
            // 这样你至少能在界面上看到它，知道抓取成功了
            if (deptName) {
                depts.add(deptName);
                
                // 仅在控制台提示配置缺失，不阻断流程
                if (!window.EMAIL_DATA || !window.EMAIL_DATA[deptName]) {
                    console.warn(`⚠️ 提示: 抓取到 "${deptName}"，但 EMAIL_DATA 中未配置对应的邮件接收人。`);
                }
            }
        } else {
            console.warn("❌ 未匹配到销售经理的部门信息，请检查铁三角文本格式。");
        }
        
        // --- 2. 获取交付列表中的部门 ---
        const tableRows = DOMElements.deliveryDetailsTable.querySelectorAll('tbody tr');
        tableRows.forEach(row => {
            const select = row.querySelector('select'); 
            if (select) {
                const dept = select.value;
                if (dept) depts.add(dept);
            }
        });

        // --- 3. 生成界面 ---
        if (depts.size === 0) { 
            addEmailDeptRow(true);
        } else {
            depts.forEach(dept => addEmailDeptRow(false, dept, true));
        }
        
        updateEmailList();
    };


    const addEmailDeptRow = (isDynamic = false, deptName = "", isChecked = false) => {
        const rowWidget = document.createElement('div');
        rowWidget.className = 'mail-dept-item'; // 应用网格布局
        
        if (isDynamic) {
            // --- 动态行 (下拉框) ---
            const control = document.createElement('select');
            control.innerHTML = [''].concat(Object.keys(EMAIL_DATA)).map(d => `<option value="${d}">${d}</option>`).join('');
            control.value = deptName;
            control.addEventListener('change', updateEmailList);
            
            rowWidget.appendChild(control); // 子元素 1: 下拉框

        } else {
            // --- 静态行 (复选框) ---
            const control = document.createElement('input');
            control.type = 'checkbox';
            control.id = `email-check-${deptName.replace(/\s+/g, '-')}`;
            control.checked = isChecked;
            control.dataset.dept = deptName;
            control.addEventListener('change', updateEmailList);

            const deptSpan = document.createElement('span');
            deptSpan.textContent = deptName;
            
            // (可选的良好体验) 点击部门名称也能选中复选框
            deptSpan.style.cursor = 'pointer';
            deptSpan.onclick = () => {
                control.checked = !control.checked;
                // 手动触发 change 事件来更新邮件列表
                control.dispatchEvent(new Event('change', { bubbles: true }));
            };

            rowWidget.appendChild(control);  // 子元素 1: 复选框
            rowWidget.appendChild(deptSpan); // 子元素 2: 部门名称
        }

        // --- 通用的移除按钮 ---
        const removeBtn = document.createElement('button');
        removeBtn.className = 'btn-remove'; // 使用您统一的 btn-remove 样式
        removeBtn.textContent = '移除';
        removeBtn.onclick = () => {
            rowWidget.remove();
            updateEmailList();
        };

        rowWidget.appendChild(removeBtn); // 子元素 3 (静态行) 或 子元素 2 (动态行)

        DOMElements.emailChecklist.appendChild(rowWidget);
    };
    
    const updateEmailList = () => {
        let leaders = [], managers = [], seenEmails = new Set();
        
        // --- 关键修改：查找 .mail-dept-item 而不是 .checklist-row ---
        const rows = DOMElements.emailChecklist.querySelectorAll('.mail-dept-item');
        // --- 修改结束 ---
        
        const getDeptFromRow = row => {
            const chk = row.querySelector('input[type="checkbox"]');
            if (chk && chk.checked) return chk.dataset.dept;
            const sel = row.querySelector('select');
            if (sel && sel.value) return sel.value;
            return null;
        };

        rows.forEach(row => {
            const dept = getDeptFromRow(row);
            if (dept && EMAIL_DATA[dept]) {
                const [name, email] = EMAIL_DATA[dept].leader;
                if (email && !seenEmails.has(email)) {
                    leaders.push(`${name} <${email}>`);
                    seenEmails.add(email);
                }
            }
        });
        rows.forEach(row => {
            const dept = getDeptFromRow(row);
            if (dept && EMAIL_DATA[dept]) {
                const [name, email] = EMAIL_DATA[dept].manager;
                if (email && !seenEmails.has(email)) {
                    managers.push(`${name} <${email}>`);
                    seenEmails.add(email);
                }
            }
        });
        
        let finalText = leaders.join('； ');
        if (leaders.length > 0 && managers.length > 0) finalText += '；\n';
        finalText += managers.join('； ');
        DOMElements.emailOutput.value = finalText;
    };
    
    document.getElementById('email-add-btn').addEventListener('click', () => addEmailDeptRow(true));
    document.getElementById('email-refresh-btn').addEventListener('click', repopulateEmailDepartments);

    // ======================= IMPORT/EXPORT LOGIC ======================
    
    document.getElementById('btn-export').addEventListener('click', () => {
        const data = gatherFormData();

        // 【【【【【【 新增的唯一一行代码 】】】】】】
        saveDataToBackend(data); // 在导出本地文件的同时，发送到服务器
        // 【【【【【【        结束         】】】】】】


        const projectName = data.projectName || '纪要数据';
        const businessCode = data.businessCode ? `(${data.businessCode})` : '';
        
        // --- 修复时区问题的代码 ---
        const now = new Date(); // 获取当前本地时间
        const Y = now.getFullYear();
        const M = (now.getMonth() + 1).toString().padStart(2, '0'); // 月份从0开始，所以+1
        const D = now.getDate().toString().padStart(2, '0');
        const h = now.getHours().toString().padStart(2, '0');
        const m = now.getMinutes().toString().padStart(2, '0');
        const timestamp = `${Y}${M}${D}${h}${m}`; // 拼接为 YYYYMMDDHHMM 格式
        // --- 修复结束 ---
        
        const filename = `${projectName}${businessCode}_${timestamp}.json`;
        const blob = new Blob([JSON.stringify(data, null, 4)], { type: 'application/json' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename.replace(/[<>:"/\\|?*]/g, '_');;
        link.click();
        URL.revokeObjectURL(link.href);
    });

    const importInput = document.getElementById('import-file-input');
    document.getElementById('btn-import').addEventListener('click', () => importInput.click());
    importInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = JSON.parse(e.target.result);
                let importData = data;
                if (Array.isArray(data) && data.length > 0) {
                    console.log('检测到数据库导出格式 (数组)，将导入第一个条目。');
                    importData = data[0]; // 只获取数组中的第一个对象
            }

                
                populateFormData(importData);
                alert('数据导入成功！');
            } catch (err) { alert(`导入失败: ${err.message}`); }
        };
        reader.readAsText(file);
        importInput.value = '';
    });
    
    function populateFormData(data) {
        const inputs = DOMElements.mainForm.querySelectorAll('input, select, textarea');
        inputs.forEach(input => {
            if (input.id && data.hasOwnProperty(input.id)) {
                input.value = data[input.id];
                input.dispatchEvent(new Event('change', { bubbles: true }));
            }
        });

        if (data.deliveryDetails) {
            const rows = DOMElements.deliveryDetailsTable.querySelector('tbody').rows;
            for(let i=0; i < rows.length; i++){ // Clear existing data first
                rows[i].cells[1].querySelector('select').value = '';
                rows[i].cells[2].querySelector('input').value = '';
                rows[i].cells[3].querySelector('input').value = '';
                rows[i].cells[4].querySelector('input').value = '';
            }
            data.deliveryDetails.forEach((rowData, index) => {
                if (index < rows.length) {
                    const row = rows[index];
                    row.cells[1].querySelector('select').value = rowData['事业部'] || '';
                    row.cells[2].querySelector('input').value = rowData['项目经理'] || '';
                    row.cells[3].querySelector('input').value = rowData['交付内容'] || '';
                    row.cells[4].querySelector('input').value = rowData['预算（万元）'] || '';
                }
            });
        }
    }
    
    // ======================= 3. 修复“返回目录”按钮导致左侧跳动 =======================
    const backToTocBtn = document.querySelector('.help-title-button');
    
    if (backToTocBtn) {
        backToTocBtn.addEventListener('click', function(e) {
            e.preventDefault(); // 1. 踩刹车：阻止浏览器默认的整页滚动
            
            // 2. 手动让右侧帮助容器滚动到最顶部 (0的位置)
            const container = document.querySelector('.help-content');
            if (container) {
                container.scrollTo({
                    top: 0, 
                    behavior: 'smooth'
                });
            }
        });
    }











    
/////////////////////////////////////////////////////////////////////////////////////
    // ======================= 免填写粘贴 =======================
    // 绑定按钮事件
    const pasteBtn = document.getElementById('btn-smart-paste');
    if (pasteBtn) {
        pasteBtn.addEventListener('click', handleSmartPaste);
    }
});

    // ======================= 最终稳健版 (修复预算漏抓 + 商机编码串标) =======================
    async function handleSmartPaste() {
        try {
            let text = '';
            if (navigator.clipboard && navigator.clipboard.readText) {
                try { text = await navigator.clipboard.readText(); } catch (e) {}
            }
            if (!text || !text.trim()) {
                const manualText = prompt("⚠️ 请在此框中按下 Ctrl+V 粘贴内容：");
                if (manualText) text = manualText;
            }
            if (!text || !text.trim()) return;
        
            // 1. 解析数据
            const data = parseVerticalData(text);
            console.log("解析结果:", data); 
        
            let count = 0;
        
            // --- A. 基础字段 ---
            count += setInputValue('projectName', data.projectName);
            
            // [修复] 商机编码：遇到空格/换行/“合同”二字立即截断
            if (data.businessCode) {
                let cleanCode = data.businessCode.split(/[\s\t\n]+|合同编号|项目类型/)[0];
                count += setInputValue('businessCode', cleanCode);
            }
        
            count += setInputValue('contractClient', data.client);
            count += setInputValue('constructionContent', data.content);
            
            // --- B. [重点修复] 预算金额 ---
            if (data.budget) {
                // 提取金额数字（支持逗号）
                const match = data.budget.match(/[\d,]+(\.\d+)?/);
                if (match) {
                    const money = match[0].replace(/,/g, '');
                    count += setInputValue('budgetAmount', money);
                }
            }
        
            // --- C. 下拉框 ---
            count += setSelectValue('projectLevel', data.level);
            count += setSelectValue('capacityType', data.capacityType);
            
            // --- D. 外采联动 (修复布局显示) ---
            if (data.extBudget) {
                const match = data.extBudget.match(/[\d,]+(\.\d+)?/);
                if (match) {
                    let extMoney = parseFloat(match[0].replace(/,/g, ''));
                    if (extMoney > 0) {
                        count += setSelectValue('procurement', '是');
                        const procInput = document.getElementById('procurementAmount');
                        const coreSelect = document.getElementById('coreCapability');
                        if (procInput) procInput.disabled = false; 
                        if (coreSelect) coreSelect.disabled = false;
                        count += setInputValue('procurementAmount', extMoney);
                        
                        if (data.procurementSituation) {
                            const riskRow = document.getElementById('procurementRiskRow');
                            // 使用空字符串恢复默认 display (flex/block)
                            if (riskRow) riskRow.style.display = ''; 
                            count += setInputValue('procurementRisk', data.procurementSituation);
                        }
                    } else {
                        setSelectValue('procurement', '否');
                    }
                }
            } else if (data.procurement) {
                count += setSelectValue('procurement', data.procurement);
            }
        
            // --- E. 毛利率/净利率 ---
            if (data.grossMargin) {
                const match = data.grossMargin.match(/-?\d+(\.\d+)?/);
                if (match) {
                    count += setInputValue('SJ_grossMargin', match[0]);
                    count += setInputValue('TB_grossMargin', match[0]);
                }
            }
            if (data.netMargin) {
                const match = data.netMargin.match(/-?\d+(\.\d+)?/);
                if (match) count += setInputValue('SJ_netMargin', match[0]);
            }
            
            // --- F. 其他字段 ---
            if (data.biddingMethod) {
                count += setSelectFuzzy('TB_biddingMethod', data.biddingMethod);
                count += setSelectFuzzy('JD_biddingMethod', data.biddingMethod);
            }
            count += setSelectValue('TB_biddingEntity', data.biddingEntity);
            if (data.businessType) {
                count += setSelectValue('JD_businessType', data.businessType);
                count += setSelectValue('TB_businessType', data.businessType);
            }
            if (data.duration) {
                count += setInputValue('JD_deliveryPeriod', data.duration);
                count += setInputValue('TB_deliveryPeriod', data.duration);
            }
        
            // --- G. 铁三角 ---
            if (data.pm || data.sales || data.solution || data.delivery) {
                const ironText = `项目经理：${data.pm || ''}\n销售经理：${data.sales || ''}\n方案经理：${data.solution || ''}\n交付经理：${data.delivery || ''}`;
                const ironInput = document.getElementById('ironTriangleInput');
                if (ironInput) {
                    ironInput.value = ironText;
                    highlightInput(ironInput);
                    count++;
                }
            }
        
            // 反馈
            if (count > 0) {
                const btn = document.getElementById('btn-smart-paste');
                if (btn) {
                    const originalText = btn.innerText;
                    btn.innerText = `✅ 成功填充 ${count} 项`;
                    btn.style.backgroundColor = '#28a745';
                    setTimeout(() => {
                        btn.innerText = originalText;
                        btn.style.backgroundColor = '';
                    }, 2000);
                } 
            } else {
                alert('未识别到有效数据，请检查复制内容。');
            }
        
        } catch (err) {
            console.error('粘贴出错:', err);
            alert('程序错误，请查看控制台。');
        }
    }
    
    // --- 核心解析器 ---
    function parseVerticalData(text) {
        const result = {};
        const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l);
    
        const isKey = (str) => {
            if (str.endsWith('：') || str.endsWith(':')) return true;
            if (str === '请选择' || str.startsWith('请选择')) return true; 
            const keys = [
                '项目预算', '项目名称', '商机编号', '合同编号', '基本信息', '会前信息', 
                '会议内容记录', '问题及解决方案', '会议决议', '参会人员', '铁三角', 
                '项目后向采购是否', '核心能力标签', '项目各板块需求', '项目实施可行性',
                '技术要求', '总体方案', '招标文件', '运维服务要求', '应急方案', 
                '是否包含监控', '请确认项目类型', '是否需要签订', '项目外采评估', 
                '人工成本评估', '列收方式', '外采评估', '是否需要标前引入', 
                '项目后向采购基本情况'
            ];
            return keys.some(k => str.startsWith(k));
        };
    
        const isMoneyLine = (str) => {
            if (!/\d/.test(str)) return false; 
            if (str.includes('结构') || str.includes('含税') || str.includes('不含税')) return false; 
            // 排除纯文字标签，但如果标签后跟着数字则保留
            if (str.includes('合同金额') && !/\d/.test(str.replace('合同金额',''))) return false; 
            if ((str.match(/-/g)||[]).length >= 2 || (str.match(/\//g)||[]).length >= 2) return false;
            return true;
        };
    
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            
            const getValue = () => {
                const colParts = line.split(/[:：]/);
                if (colParts.length > 1 && colParts[1].trim()) return colParts[1].trim();
                if (i + 1 < lines.length) return lines[i+1];
                return '';
            };
        
            // 1. 商机编号
            if (line.includes('商机编号') && !line.includes('mss')) {
                result.businessCode = getValue(); 
            }
        
            // 2. [重点修复] 项目预算
            // 必须以 "项目预算" 开头，防止匹配到 "项目级别...含税的项目预算" 这种描述文字
            else if (line.startsWith('项目预算')) {
                const parts = line.split(/[:：]/);
                // 2.1 先看当前行冒号后有没有钱
                if (parts.length > 1 && isMoneyLine(parts[1])) {
                    result.budget = parts[1];
                } else {
                    // 2.2 否则往后找8行 (加大范围，跳过所有干扰)
                    for (let k = 1; k <= 8; k++) {
                        if (i + k >= lines.length) break;
                        const nextRow = lines[i + k];
                        
                        // 遇到明显的新大标题才停，但跳过 "合同金额" 这种伪标题
                        if (isKey(nextRow) && !nextRow.includes('合同金额') && !nextRow.includes('软件金额')) break;
                        
                        if (isMoneyLine(nextRow)) {
                            result.budget = nextRow;
                            break; 
                        }
                    }
                }
            }
        
            // 3. 毛利率/净利率
            else if (line.includes('毛利率')) {
                let match = line.match(/(-?\d+(\.\d+)?)%/);
                result.grossMargin = match ? match[0] : getValue();
            }
            else if (line.includes('净利润率')) {
                let match = line.match(/(-?\d+(\.\d+)?)%/);
                result.netMargin = match ? match[0] : getValue();
            }
        
            // 4. 建设内容 (多行)
            else if (line.startsWith('建设内容')) {
                let contentArr = [];
                const parts = line.split(/[:：]/);
                if (parts.length > 1 && parts[1].trim()) contentArr.push(parts[1].trim());
                let j = i + 1;
                while (j < lines.length) {
                    const nextRow = lines[j];
                    // 停止条件：遇到 Key (且不以数字开头)
                    if (nextRow.startsWith('项目预算') || (isKey(nextRow) && !/^\d+[、\.]/.test(nextRow))) break;
                    contentArr.push(nextRow);
                    j++;
                }
                result.content = contentArr.join('\n');
                i = j - 1;
            }
        
            // 5. 外采风险 (多行 + 截断)
            else if (line.startsWith('项目后向采购基本情况')) {
                let contentArr = [];
                const parts = line.split(/[:：]/);
                if (parts.length > 1 && parts[1].trim()) contentArr.push(parts[1].trim());
                let j = i + 1;
                while (j < lines.length) {
                    const nextRow = lines[j];
                    if (isKey(nextRow) && !/^\d+[、\.]/.test(nextRow)) break;
                    contentArr.push(nextRow);
                    j++;
                }
                result.procurementSituation = contentArr.join('\n');
                i = j - 1;
            }
        
            // 6. 外部采购预算
            else if (line.includes('外部采购预算') || line.includes('外部采购金额')) {
                for (let k = 1; k <= 3; k++) {
                    if (i + k >= lines.length) break;
                    const nextRow = lines[i + k];
                    if (/\d/.test(nextRow) && !nextRow.includes('结构')) {
                        result.extBudget = nextRow;
                        break;
                    }
                }
            }
        
            // 7. 其他字段
            else if (line.includes('项目名称') && !line.includes('ID')) result.projectName = getValue();
            else if (line.startsWith('签约客户') || line.startsWith('客户名称')) result.client = getValue();
            else if (line.startsWith('项目级别')) { const val = getValue(); if (val.length < 10) result.level = val; }
            else if (line.includes('产品能力')) result.capacityType = getValue();
            else if (line.includes('是否需要后向采购')) result.procurement = getValue();
            else if (line.startsWith('业务类型')) result.businessType = getValue();
            else if (line.startsWith('投标方式') || line.startsWith('签约类型')) result.biddingMethod = getValue();
            else if (line.startsWith('投标主体')) result.biddingEntity = getValue();
            else if (line.includes('项目工期') || line.includes('服务期')) result.duration = getValue();
            else if (line.startsWith('项目经理')) result.pm = getValue();
            else if (line.startsWith('销售经理')) result.sales = getValue();
            else if (line.startsWith('方案经理') || line.startsWith('售前解方经理')) result.solution = getValue();
            else if (line.startsWith('交付经理')) result.delivery = getValue();
        }
        return result;
    }
    
    // --- 辅助函数 ---
    function setInputValue(id, value) {
        if (!value) return 0;
        const el = document.getElementById(id);
        if (el) {
            el.value = value;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            highlightInput(el); return 1;
        } return 0;
    }
    function setSelectValue(id, value) {
        if (!value) return 0;
        const el = document.getElementById(id);
        if (el) {
            for (let i = 0; i < el.options.length; i++) {
                if (el.options[i].value === value || el.options[i].text === value) {
                    el.selectedIndex = i;
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    highlightInput(el); return 1;
                }
            }
        } return 0;
    }
    function setSelectFuzzy(id, textToMatch) {
        if (!textToMatch) return 0;
        const el = document.getElementById(id);
        if (el) {
            for (let i = 0; i < el.options.length; i++) {
                const optText = el.options[i].text;
                if (optText && (textToMatch.includes(optText) || optText.includes(textToMatch))) {
                    el.selectedIndex = i;
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    highlightInput(el); return 1;
                }
            }
        } return 0;
    }
    function highlightInput(element) {
        element.style.transition = 'background-color 0.3s';
        element.style.backgroundColor = '#d1e7dd';
        setTimeout(() => { element.style.backgroundColor = ''; }, 800);
    }

; // End of DOMContentLoaded



/**
 * 【新增】使用 fetch API 将数据发送到您的后端服务器
 * @param {object} data - 从 gatherFormData() 获得的对象
 */
async function saveDataToBackend(data) {
    
    // 这就是您服务器的地址！
    const backendUrl = 'https://www.gzchenjin.com/api/save';

    try {
        const response = await fetch(backendUrl, {
            method: 'POST', // 使用 POST 方法
            headers: {
                'Content-Type': 'application/json', // 告诉服务器我们发送的是 JSON
            },
            body: JSON.stringify(data), // 将 JS 对象转换为 JSON 字符串
        });

        if (response.ok) {
            // 如果服务器返回成功 (HTTP 201)
            console.log('数据成功保存到后端！');
            //alert('数据已成功备份到服务器！'); // 弹出成功提示
        } else {
            // 如果服务器返回错误 (例如 HTTP 500)
            console.error('保存到后端失败:', await response.text());
            //alert('数据备份到服务器失败，请检查服务器日志。'); // 弹出错误提示
        }
    } catch (error) {
        // 如果网络不通或服务器未运行
        console.error('连接后端服务器时出错:', error);
        //alert('无法连接到后端服务器，请检查服务器是否正在运行且防火墙已配置。'); // 弹出连接错误提示
    }
}
