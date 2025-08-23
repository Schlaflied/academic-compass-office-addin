/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// 使用 Office.onReady 作为唯一的启动器，确保 Office 环境和 DOM 都已就绪
Office.onReady((info) => {
  // 确认宿主应用是 Word
  if (info.host === Office.HostType.Word) {
    // 执行所有初始化操作
    try {
      initializeApp();
    } catch (error)      {
      console.error("初始化插件失败:", error);
    }
  }
});

/**
 * 插件的主初始化函数
 */
function initializeApp() {
  console.log("initializeApp: 函数已启动。");

  // --- 1. 获取所有需要的 DOM 元素 ---
  const majorInput = document.getElementById('major-input');
  const interestsInput = document.getElementById('interests-input');
  const resumeInput = document.getElementById('resume-input');
  const analyzeButton = document.getElementById('analyze-button');
  const resultContainer = document.getElementById('result-container');
  const sourcesContainer = document.getElementById('sources-container');
  const logo = document.getElementById('logo');
  const resizer = document.getElementById('resizer');
  const topPanel = document.getElementById('top-panel');
  
  // --- 2. 语言翻译和图标资源 (保持不变) ---
  const translations = {
    'zh-CN': {
        logo_text: '🧭 学术罗盘', title: '输入信息', subtitle: 'AI将分析你的未来可能性',
        major_label: '你的专业/学位', major_placeholder: '例如：人机交互博士',
        interests_label: '研究方向或技能 (可选)', interests_placeholder: '例如：自然语言处理',
        resume_label: '我的简历 / 个人简介 (可选)', resume_placeholder: '粘贴你的简历...',
        button_text: '开始分析', button_loading_text: '分析中...',
        result_placeholder_title: '分析报告', result_placeholder_text: '（请输入专业后点击分析）',
        sources_title: '引用来源:', support_text: '请开发者喝杯咖啡',
        theme_switch_to_dark: '切换到暗色模式', theme_switch_to_light: '切换到明亮模式',
        rate_limit_exceeded: "同学，您今日的免费探索次数已用尽！🧭\n\nAcademic Compass 每天为所有用户提供5次免费生涯规划分析。\n如果需要更多支持，欢迎明天再来探索，或通过‘请我喝杯咖啡☕️’来支持项目发展！",
        connection_error: "发生连接错误，请检查网络或联系开发者。",
        loading_statuses: [
            "正在连接AI大脑...", "正在搜索相关职业路径...", "正在分析加拿大就业市场数据...",
            "正在召唤 Gemini 进行深度分析...", "即将完成，正在生成专属生涯报告..."
        ]
    },
    'zh-TW': {
        logo_text: '🧭 學術羅盤', title: '輸入資訊', subtitle: 'AI將分析你的未來可能性',
        major_label: '你的專業/學位', major_placeholder: '例如：人機互動博士',
        interests_label: '研究方向或技能 (可選)', interests_placeholder: '例如：自然語言處理',
        resume_label: '我的履歷 / 個人簡介 (可選)', resume_placeholder: '貼上你的履歷...',
        button_text: '開始分析', button_loading_text: '分析中...',
        result_placeholder_title: '分析報告', result_placeholder_text: '（請輸入專業後點擊分析）',
        sources_title: '引用來源:', support_text: '請開發者喝杯咖啡',
        theme_switch_to_dark: '切換到暗色模式', theme_switch_to_light: '切換到明亮模式',
        rate_limit_exceeded: "同學，您今日的免費探索次數已用盡！🧭\n\nAcademic Compass 每天為所有用戶提供5次免費生涯規劃分析。\n如果需要更多支持，歡迎明天再來探索，或通過「請我喝杯咖啡☕️」來支持項目發展！",
        connection_error: "發生連接錯誤，請檢查網路或聯絡開發者。",
        loading_statuses: [
            "正在連接AI大腦...", "正在搜尋相關職業路徑...", "正在分析加拿大就業市場數據...",
            "正在召喚 Gemini进行深度分析...", "即將完成，正在生成專屬生涯報告..."
        ]
    },
    'en': {
        logo_text: '🧭 Academic Compass', title: 'Input Information', subtitle: 'AI will analyze your future possibilities',
        major_label: 'Your Major/Degree', major_placeholder: 'e.g., PhD in Human-Computer Interaction',
        interests_label: 'Research Interests or Skills (Optional)', interests_placeholder: 'e.g., Natural Language Processing',
        resume_label: 'My Resume / Bio (Optional)', resume_placeholder: 'Paste your resume...',
        button_text: 'Analyze', button_loading_text: 'Analyzing...',
        result_placeholder_title: 'Analysis Report', result_placeholder_text: '(Enter your major and click Analyze)',
        sources_title: 'References:', support_text: 'Buy the developer a coffee',
        theme_switch_to_dark: 'Switch to Dark Mode', theme_switch_to_light: 'Switch to Light Mode',
        rate_limit_exceeded: "You have used up your free explorations for today! 🧭\n\nAcademic Compass provides 5 free career analyses per day for all users.\nFeel free to come back tomorrow for more insights, or 'Buy me a coffee ☕️' to support the project!",
        connection_error: "Connection error. Please check your network or contact the developer.",
        loading_statuses: [
            "Connecting to the AI brain...", "Searching for relevant career paths...", "Analyzing Canadian job market data...",
            "Summoning Gemini for deep analysis...", "Finalizing, generating your personalized career report..."
        ]
    }
  };
  const ICONS = {
    linkedin: `<svg class="source-icon" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16" fill="currentColor"><path d="M0 1.146C0 .513.526 0 1.175 0h13.65C15.474 0 16 .513 16 1.146v13.708c0 .633-.526 1.146-1.175 1.146H1.175C.526 16 0 15.487 0 14.854V1.146zm4.943 12.248V6.169H2.542v7.225h2.401zm-1.2-8.212c.837 0 1.358-.554 1.358-1.248-.015-.709-.52-1.248-1.342-1.248-.822 0-1.359.54-1.359 1.248 0 .694.521 1.248 1.327 1.248h.016zm4.908 8.212V9.359c0-.216.016-.432.08-.586.173-.431.568-.878 1.232-.878.869 0 1.216.662 1.216 1.634v3.865h2.401V9.25c0-2.22-1.184-3.252-2.764-3.252-1.274 0-1.845.7-2.165 1.193v.025h-.016a5.54 5.54 0 0 1 .016-.025V6.169h-2.4c.03.678 0 7.225 0 7.225h2.4z"/></svg>`,
    glassdoor: `<svg class="source-icon" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16" fill="currentColor"><path fill-rule="evenodd" d="M1.185 1.185A1.5 1.5 0 0 1 2.57.293l10.854 10.854a.5.5 0 0 1 0 .708L11.146 14a.5.5 0 0 1-.708 0L.293 2.854A1.5 1.5 0 0 1 1.185 1.185zM14.815 1.185a1.5 1.5 0 0 0-2.122 0L.854 13.146a.5.5 0 0 0 0 .708L2.854 15.707a.5.5 0 0 0 .708 0L15.707 3.565a1.5 1.5 0 0 0 0-2.122l-.892-.892z"/></svg>`,
    indeed: `<svg class="source-icon" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16" fill="currentColor"><path d="M13.555 5.582a.363.363 0 0 0-.363.363v4.062a.363.363 0 0 0 .363.363h.363a.363.363 0 0 0 .363-.363V5.945a.363.363 0 0 0-.363-.363h-.363zM10.31 5.582a.363.363 0 0 0-.363.363v4.062a.363.363 0 0 0 .363.363h.363a.363.363 0 0 0 .363-.363V5.945a.363.363 0 0 0-.363-.363h-.363zM8.36 5.582a.363.363 0 0 0-.363.363v4.062a.363.363 0 0 0 .363.363h.363a.363.363 0 0 0 .363-.363V5.945a.363.363 0 0 0-.363-.363h-.363zM5.945 5.582a.363.363 0 0 0-.363.363v4.062a.363.363 0 0 0 .363.363h.363a.363.363 0 0 0 .363-.363V5.945a.363.363 0 0 0-.363-.363h-.363zM15.363 4.091A1.91 1.91 0 0 0 13.455 2.182h-10.91A1.91 1.91 0 0 0 .636 4.091v7.818A1.91 1.91 0 0 0 2.545 13.818h10.91a1.91 1.91 0 0 0 1.909-1.909V4.091zM2.909 5.227a1.136 1.136 0 1 1 0 2.273 1.136 1.136 0 0 1 0-2.273z"/></svg>`,
    default: `<svg class="source-icon" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16" fill="currentColor"><path d="M4.715 6.542 3.343 7.914a3 3 0 1 0 4.243 4.243l1.828-1.829A3 3 0 0 0 8.586 5.5L8 6.086a1.002 1.002 0 0 0-.154.199 2 2 0 0 1 .861 3.337L6.88 11.45a2 2 0 1 1-2.83-2.83l.793-.792a4.018 4.018 0 0 1-.128-1.287z"/><path d="M6.586 4.672A3 3 0 0 0 7.414 9.5l.775-.776a2 2 0 0 1-.896-3.346L9.12 3.55a2 2 0 1 1 2.83 2.83l-.793.792c.112.42.155.855.128 1.287l1.372-1.372a3 3 0 1 0-4.243-4.243L6.586 4.672z"/></svg>`
  };

  // --- 3. 核心功能逻辑 ---
  let currentLang = 'zh-CN';
  let loadingInterval = null;
  const API_URL = 'https://academic-compass-backend-885033581194.us-central1.run.app/analyze'; 

  function applyLanguage(langCode) {
    currentLang = langCode;
    const t = translations[langCode] || translations['en'];
    document.querySelectorAll('[data-key]').forEach(elem => { const key = elem.getAttribute('data-key'); if (t[key]) elem.textContent = t[key]; });
    document.querySelectorAll('[data-key-placeholder]').forEach(elem => { const key = elem.getAttribute('data-key-placeholder'); if (t[key]) elem.placeholder = t[key]; });
    logo.textContent = t.logo_text;
    document.getElementById('lang-toggle').querySelectorAll('button').forEach(button => {
        button.classList.toggle('active', button.dataset.lang === langCode);
    });
  }

  // 【v9 修正】应用主题的函数，增加了强制渲染的技巧
  function applyTheme(theme) {
    const t = translations[currentLang] || translations['en'];
    const isLight = theme === 'light';
    const themeSwitcher = document.getElementById('theme-switcher');
    
    // 步骤1: 改变整个文档的 class，这个通常不会有渲染问题
    if (isLight) {
        document.documentElement.classList.add('light-mode');
    } else {
        document.documentElement.classList.remove('light-mode');
    }
    
    // 步骤2: 【核心魔法】使用 setTimeout(..., 0) 来强制渲染引擎重绘按钮
    if(themeSwitcher) {
        setTimeout(() => {
            themeSwitcher.textContent = isLight ? '🌙' : '☀️';
            themeSwitcher.title = isLight ? t.theme_switch_to_dark : t.theme_switch_to_light;
        }, 0);
    }
  }
  
  // 保存设置的函数，增加了 localStorage 作为备用方案
  function saveSettings(key, value) {
      // 方案A: 尝试使用 Office 的原生方式保存
      try {
          Office.context.document.settings.set(key, value);
          Office.context.document.settings.saveAsync((asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error("Office.context.document.settings.saveAsync() 失败: ", asyncResult.error.message);
              } else {
                  console.log(`通过 Office API 成功保存设置: ${key} = ${value}`);
              }
          });
      } catch (e) {
          console.error("调用 Office.context.document.settings.set() 时出错: ", e);
      }

      // 方案B: 无论方案A是否成功，都使用 localStorage 保存一份
      try {
          localStorage.setItem(key, value);
          console.log(`通过 localStorage 成功保存设置: ${key} = ${value}`);
      } catch (e) {
          console.error("调用 localStorage.setItem() 时出错: ", e);
      }
  }

  // 加载设置的函数，增加了 localStorage 作为备用方案
  function loadSettings(key, defaultValue) {
      // 优先从 localStorage 读取，因为它更快且通常不会被阻止
      const localValue = localStorage.getItem(key);
      if (localValue !== null) {
          console.log(`从 localStorage 加载到设置: ${key} = ${localValue}`);
          return localValue;
      }
      
      // 如果 localStorage 中没有，再尝试从 Office 设置中读取
      const officeValue = Office.context.document.settings.get(key);
      if (officeValue !== null && officeValue !== undefined) {
          console.log(`从 Office API 加载到设置: ${key} = ${officeValue}`);
          return officeValue;
      }
      
      // 如果都没有，则返回默认值
      console.log(`未找到 '${key}' 的任何已保存设置，使用默认值: ${defaultValue}`);
      return defaultValue;
  }


  analyzeButton.addEventListener('click', async () => {
    const t = translations[currentLang];
    const buttonTextSpan = analyzeButton.querySelector('span');
    const existingSpinner = analyzeButton.querySelector('.spinner');
    if (existingSpinner) { existingSpinner.remove(); }
    buttonTextSpan.textContent = t.button_loading_text;
    analyzeButton.insertAdjacentHTML('beforeend', '<div class="spinner"></div>');
    analyzeButton.disabled = true;
    sourcesContainer.innerHTML = '';
    let statusIndex = 0;
    const loadingStatuses = t.loading_statuses;
    resultContainer.innerHTML = `<h2 data-key="result_placeholder_title">${t.result_placeholder_title}</h2><p>${loadingStatuses[statusIndex]}</p>`;
    statusIndex++;
    loadingInterval = setInterval(() => {
        if (statusIndex < loadingStatuses.length) {
            resultContainer.innerHTML = `<h2 data-key="result_placeholder_title">${t.result_placeholder_title}</h2><p>${loadingStatuses[statusIndex]}</p>`;
            statusIndex++;
        } else {
            clearInterval(loadingInterval);
        }
    }, 2500);
    const analysisData = { major: majorInput.value, interests: interestsInput.value, resumeText: resumeInput.value, language: currentLang };
    try {
        const response = await fetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(analysisData) });
        const result = await response.json();
        if (loadingInterval) clearInterval(loadingInterval);
        const heading = `<h2 data-key="result_placeholder_title">${t.result_placeholder_title}</h2>`;
        if (response.ok) {
            let analysisHtml = marked.parse(result.analysis || '');
            analysisHtml = analysisHtml.replace(/\[(\d+)\]/g, (match, number) => `<a href="#source-${number}" class="citation-link">${match}</a>`);
            resultContainer.innerHTML = DOMPurify.sanitize(heading + analysisHtml);
            sourcesContainer.innerHTML = ''; 
            if (result.sources && result.sources.length > 0) {
                let sourcesHTML = `<h2>${t.sources_title}</h2>`;
                result.sources.forEach(source => {
                    const icon = ICONS[source.source_type] || ICONS.default;
                    sourcesHTML += `<div class="source-item" id="source-${source.id}">${icon}<span>[${source.id}] <a href="${source.link}" target="_blank">${source.title}</a></span></div>`;
                });
                sourcesContainer.innerHTML = DOMPurify.sanitize(sourcesHTML, {ADD_ATTR: ['id'], ADD_TAGS: ['svg', 'path']});
            }
        } else {
            if (result.error === 'rate_limit_exceeded') {
                resultContainer.innerHTML = `${heading}<p style="white-space: pre-wrap;">${result.message || t.rate_limit_exceeded}</p>`;
            } else {
                resultContainer.innerHTML = `${heading}<p>${result.error || 'Unknown error'}</p>`;
            }
        }
    } catch (error) {
        if (loadingInterval) clearInterval(loadingInterval);
        resultContainer.innerHTML = `<h2>Error</h2><p>${t.connection_error}</p>`;
        console.error("Fetch Error:", error);
    } finally {
        buttonTextSpan.textContent = t.button_text;
        const finalSpinner = analyzeButton.querySelector('.spinner');
        if (finalSpinner) finalSpinner.remove();
        analyzeButton.disabled = false;
    }
  });

  function makeResizable() {
    let isResizing = false;
    resizer.addEventListener('mousedown', (e) => { isResizing = true; document.body.style.userSelect = 'none'; });
    window.addEventListener('mousemove', (e) => {
        if (!isResizing) return;
        const newHeight = e.clientY - topPanel.offsetTop;
        const minHeight = 250;
        const maxHeight = window.innerHeight - 200;
        if (newHeight > minHeight && newHeight < maxHeight) { topPanel.style.height = `${newHeight}px`; }
    });
    window.addEventListener('mouseup', (e) => { 
        if (isResizing) { isResizing = false; document.body.style.userSelect = ''; saveSettings('ac-panel-height', topPanel.style.height); }
    });
  }

  // --- 5. 使用事件委托来绑定所有交互事件 ---
  document.body.addEventListener('click', (event) => {
    const target = event.target;

    // 检查是否点击了“切换主题”按钮
    const themeSwitcherButton = target.closest('#theme-switcher');
    if (themeSwitcherButton) {
        const isLight = document.documentElement.classList.contains('light-mode');
        const newTheme = isLight ? 'dark' : 'light';
        applyTheme(newTheme);
        saveSettings('ac-theme', newTheme);
        return;
    }

    // 检查是否点击了“切换语言”按钮
    const langButton = target.closest('.lang-button');
    if (langButton) {
        const langCode = langButton.dataset.lang;
        if (langCode && langCode !== currentLang) {
            applyLanguage(langCode);
            saveSettings('ac-language', langCode);
        }
        return;
    }

    // 检查是否点击了“收起”按钮
    const collapseBtn = target.closest('#collapse-button');
    if (collapseBtn) {
        console.log("“收起”按钮被点击，但由于 Office API 限制，无法关闭任务窗格。");
        return;
    }
  });
  
  // 启用面板拖动调整大小的功能
  makeResizable();

  // --- 6. 加载并应用已保存的设置 ---
  const savedTheme = loadSettings('ac-theme', 'dark');
  const savedLang = loadSettings('ac-language', 'zh-CN');
  const savedHeight = loadSettings('ac-panel-height', null);

  applyTheme(savedTheme);
  applyLanguage(savedLang);
  
  if (savedHeight) {
      topPanel.style.height = savedHeight;
  }
  console.log("已成功加载并应用所有保存的设置。");
}
