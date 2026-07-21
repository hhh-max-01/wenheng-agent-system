const form = document.querySelector('#judgeForm');
const filesInput = document.querySelector('#files');
const fileList = document.querySelector('#fileList');
const dropzone = document.querySelector('#dropzone');
const intent = document.querySelector('#intent');
const intentCount = document.querySelector('#intentCount');
const runButton = document.querySelector('#runButton');
const emptyState = document.querySelector('#emptyState');
const loadingState = document.querySelector('#loadingState');
const errorState = document.querySelector('#errorState');
const resultsRoot = document.querySelector('#results');
const resultSummary = document.querySelector('#resultSummary');
const copyButton = document.querySelector('#copyButton');
const modeBadge = document.querySelector('#modeBadge');
let selectedFiles = [];
let lastPayload = null;

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]));
}

function updateIntentCount() {
  intent.value = intent.value.slice(0, 200);
  intentCount.textContent = `${intent.value.length} / 200`;
}

function updateFiles(files) {
  const allowed = ['doc', 'docx', 'pdf', 'json', 'csv', 'txt', 'md'];
  selectedFiles = [...files].filter(file => allowed.includes(file.name.split('.').pop().toLowerCase()));
  fileList.innerHTML = selectedFiles.map(file => `<li><span>${escapeHtml(file.name)}</span><b>${(file.size / 1024).toFixed(1)} KB</b></li>`).join('');
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function showState(name) {
  emptyState.hidden = name !== 'empty';
  loadingState.hidden = name !== 'loading';
  errorState.hidden = name !== 'error';
  resultSummary.hidden = name !== 'results';
  resultsRoot.hidden = name !== 'results';
  copyButton.hidden = name !== 'results';
}

function renderResults(payload) {
  lastPayload = payload;
  const pass = payload.results.filter(item => item.label === '通过').length;
  const fail = payload.results.length - pass;
  resultSummary.innerHTML = `<span>共审核 <b>${payload.count}</b> 条</span><span>通过 <b>${pass}</b> · 不通过 <b>${fail}</b></span>`;
  resultsRoot.innerHTML = payload.results.map(item => {
    const rules = item.matched_rules.length
      ? item.matched_rules.map(rule => `<div class="rule"><b>${escapeHtml(rule.rule_id)} · ${escapeHtml(rule.rule_name)}</b><p>${escapeHtml(rule.evidence)}</p></div>`).join('')
      : '<div class="no-rule">未发现有充分证据支持的否决规则</div>';
    const warning = item.warning ? `<div class="no-rule">处理提示：${escapeHtml(item.warning)}</div>` : '';
    return `<article class="result-card">
      <div class="result-head"><div><h3>${escapeHtml(item.id)}</h3><small>${escapeHtml(item.source_name)} · ${escapeHtml(item.mode)}</small></div><span class="label ${item.label === '通过' ? 'pass' : 'fail'}">${escapeHtml(item.label)}</span></div>
      <div class="result-body"><p class="reason">${escapeHtml(item.reason)}</p><div class="confidence">置信度 ${(Number(item.confidence) * 100).toFixed(0)}%</div>${warning}${rules}</div>
    </article>`;
  }).join('');
  showState('results');
}

async function refreshHealth() {
  try {
    const response = await fetch('/api/health');
    const data = await response.json();
    if (data.mode === 'llm') {
      modeBadge.textContent = `${data.provider} · ${data.model}`;
      modeBadge.classList.remove('demo');
    } else {
      modeBadge.textContent = '本地规则演示模式';
      modeBadge.classList.add('demo');
    }
  } catch {
    modeBadge.textContent = '服务连接异常';
    modeBadge.classList.add('demo');
  }
}

intent.addEventListener('input', updateIntentCount);
document.querySelectorAll('[data-intent]').forEach(button => button.addEventListener('click', () => {
  intent.value = button.dataset.intent;
  updateIntentCount();
}));
filesInput.addEventListener('change', () => updateFiles(filesInput.files));
['dragenter', 'dragover'].forEach(name => dropzone.addEventListener(name, event => { event.preventDefault(); dropzone.classList.add('dragging'); }));
['dragleave', 'drop'].forEach(name => dropzone.addEventListener(name, event => { event.preventDefault(); dropzone.classList.remove('dragging'); }));
dropzone.addEventListener('drop', event => updateFiles(event.dataTransfer.files));
copyButton.addEventListener('click', async () => {
  if (!lastPayload) return;
  await navigator.clipboard.writeText(JSON.stringify(lastPayload.results, null, 2));
  copyButton.textContent = '已复制';
  setTimeout(() => copyButton.textContent = '复制 JSON', 1200);
});

form.addEventListener('submit', async event => {
  event.preventDefault();
  errorState.textContent = '';
  if (!selectedFiles.length) {
    errorState.textContent = '请先选择至少一个支持的文件。';
    showState('error');
    return;
  }
  const totalSize = selectedFiles.reduce((sum, file) => sum + file.size, 0);
  if (totalSize > 18 * 1024 * 1024) {
    errorState.textContent = '文件总量超过18MB，请减少文件数量或压缩后重试。';
    showState('error');
    return;
  }
  runButton.disabled = true;
  showState('loading');
  try {
    const files = await Promise.all(selectedFiles.map(async file => ({name: file.name, type: file.type, base64: await fileToBase64(file)})));
    const response = await fetch('/api/analyze', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        dataset_type: document.querySelector('input[name="datasetType"]:checked').value,
        intent: intent.value.trim(),
        files,
      }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || '审核请求失败');
    renderResults(payload);
  } catch (error) {
    errorState.textContent = error.message || '系统暂时无法完成审核，请稍后重试。';
    showState('error');
  } finally {
    runButton.disabled = false;
  }
});

updateIntentCount();
refreshHealth();
