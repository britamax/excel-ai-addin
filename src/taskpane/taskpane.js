/* global Office, Excel */
(function () {
  'use strict';

  const STORAGE_KEY = 'aiexcel_config';
  let config = { apiKey: '', model: '', endpoint: '', temperature: 0.7 };
  let chatHistory = [];
  let abortController = null;

  // --- Config ---
  function loadConfig() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) Object.assign(config, JSON.parse(raw));
    } catch (_) {}
    document.getElementById('cfg-endpoint').value = config.endpoint || '';
    document.getElementById('cfg-apikey').value = config.apiKey || '';
    document.getElementById('cfg-model').value = config.model || '';
    document.getElementById('cfg-temp').value = config.temperature || 0.7;
    document.getElementById('cfg-temp-val').textContent = config.temperature || 0.7;
  }

  function saveConfig() {
    config.endpoint = document.getElementById('cfg-endpoint').value.trim();
    config.apiKey = document.getElementById('cfg-apikey').value.trim();
    config.model = document.getElementById('cfg-model').value.trim();
    config.temperature = parseFloat(document.getElementById('cfg-temp').value) || 0.7;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
    document.getElementById('config-panel').classList.add('hidden');
    setStatus('Config saved', '');
  }

  // --- UI ---
  function addMessage(role, text) {
    const container = document.getElementById('chat-messages');
    const div = document.createElement('div');
    div.className = 'msg ' + role;
    div.innerHTML = '<div class="msg-content">' + escapeHtml(text) + '</div>';
    const typing = container.querySelector('.typing');
    if (typing && role !== 'ai') typing.remove();
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
    return div.querySelector('.msg-content');
  }

  function setTypingIndicator() {
    const container = document.getElementById('chat-messages');
    const old = container.querySelector('.typing');
    if (old) old.remove();
    const div = document.createElement('div');
    div.className = 'msg ai typing';
    div.innerHTML = '<div class="msg-content"></div>';
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
  }

  function updateMessage(el, text) {
    if (el) el.innerHTML = escapeHtml(text) + getWriteButton(text);
    document.getElementById('chat-messages').scrollTop = document.getElementById('chat-messages').scrollHeight;
  }

  function setStatus(text, type) {
    const bar = document.getElementById('status-bar');
    bar.textContent = text;
    bar.className = type ? 'visible ' + type : 'visible';
    if (!text) bar.classList.add('hidden');
  }

  function escapeHtml(s) {
    const d = document.createElement('div');
    d.textContent = s;
    return d.innerHTML;
  }

  function getWriteButton(text) {
    return '<br><button class="write-btn" onclick="window.writeToCell(\'' + escapeHtml(text.replace(/'/g, "\\'")) + '\')">📝 Write to cell</button>';
  }

  // --- API ---
  async function sendChat(userText) {
    if (!config.endpoint || !config.apiKey) {
      setStatus('Set Endpoint + API Key in ⚙️ Settings', 'error');
      document.getElementById('config-btn').click();
      return;
    }
    if (!userText.trim()) return;

    const input = document.getElementById('chat-input');
    input.value = '';
    input.style.height = 'auto';

    addMessage('user', userText);
    chatHistory.push({ role: 'user', content: userText });
    setTypingIndicator();

    const msgEl = document.querySelector('.typing .msg-content');
    let buffer = '';

    abortController = new AbortController();
    const btn = document.getElementById('send-btn');
    btn.disabled = true;

    try {
      const history = chatHistory.slice(-10);
      const body = { messages: [{ role: 'system', content: 'You are a helpful AI assistant inside Microsoft Excel. Respond concisely and accurately.' }, ...history], temperature: config.temperature, stream: true };
      if (config.model) body.model = config.model;

      const res = await fetch(config.endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + config.apiKey },
        body: JSON.stringify(body),
        signal: abortController.signal,
      });

      if (!res.ok) {
        const errText = await res.text();
        throw new Error('HTTP ' + res.status + ': ' + errText.slice(0, 200));
      }

      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let done = false;
      let fullContent = '';

      while (!done) {
        const { value, done: d } = await reader.read();
        done = d;
        if (value) {
          const chunk = decoder.decode(value, { stream: true });
          const lines = chunk.split('\n');
          for (const line of lines) {
            if (line.startsWith('data: ')) {
              const data = line.slice(6).trim();
              if (data === '[DONE]') { done = true; break; }
              try {
                const parsed = JSON.parse(data);
                const delta = parsed.choices?.[0]?.delta?.content || parsed.choices?.[0]?.delta?.reasoning_content || '';
                if (delta) { fullContent += delta; updateMessage(msgEl, fullContent); }
              } catch (_) {}
            }
          }
        }
      }

      document.querySelector('.typing')?.classList.remove('typing');
      chatHistory.push({ role: 'assistant', content: fullContent || '(empty response)' });
    } catch (err) {
      document.querySelector('.typing')?.classList.remove('typing');
      if (err.name === 'AbortError') {
        addMessage('ai', '⏹ Stopped.');
      } else {
        addMessage('ai', '❌ Error: ' + err.message);
        setStatus('API error', 'error');
      }
    } finally {
      btn.disabled = false;
      abortController = null;
    }
  }

  // --- Write to Cell ---
  window.writeToCell = function (text) {
    try {
      Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.values = [[text]];
        await context.sync();
        setStatus('✅ Written to ' + range.address, '');
      }).catch((err) => {
        navigator.clipboard.writeText(text).then(() => {
          setStatus('📋 Copied to clipboard', '');
        }).catch(() => {});
        setStatus('⚠ Could not write to cell: ' + err.message, 'error');
      });
    } catch (_) {
      navigator.clipboard.writeText(text).then(() => setStatus('📋 Copied to clipboard', ''));
    }
  };

  // --- Init ---
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel || info.platform === Office.PlatformType.OfficeOnline) {
      loadConfig();

      document.getElementById('send-btn').addEventListener('click', () => {
        sendChat(document.getElementById('chat-input').value);
      });
      document.getElementById('chat-input').addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          sendChat(e.target.value);
        }
      });
      document.getElementById('chat-input').addEventListener('input', function () {
        this.style.height = 'auto';
        this.style.height = Math.min(this.scrollHeight, 120) + 'px';
      });

      document.getElementById('config-btn').addEventListener('click', () => {
        document.getElementById('config-panel').classList.toggle('hidden');
      });
      document.getElementById('config-close').addEventListener('click', () => {
        document.getElementById('config-panel').classList.add('hidden');
      });
      document.getElementById('cfg-save').addEventListener('click', saveConfig);
      document.getElementById('cfg-temp').addEventListener('input', function () {
        document.getElementById('cfg-temp-val').textContent = this.value;
      });

      // Prompt config if empty
      if (!config.endpoint || !config.apiKey) document.getElementById('config-btn').click();
    }
  });
})();
