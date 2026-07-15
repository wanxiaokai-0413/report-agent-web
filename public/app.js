document.addEventListener("DOMContentLoaded", function () {
  const STORAGE_KEY = "report-agent-web.conversations";

  const usageModal = document.getElementById("usageModal");
  const closeUsageTip = document.getElementById("closeUsageTip");
  const openUsageTip = document.getElementById("openUsageTip");
  const fileUpload = document.getElementById("fileUpload");
  const fileList = document.getElementById("fileList");
  const reportInput = document.getElementById("reportInput");
  const sendBtn = document.getElementById("sendBtn");
  const newChatBtn = document.getElementById("newChatBtn");
  const chatArea = document.getElementById("chatArea");
  const mainContent = document.querySelector(".main-content");
  const historyList = document.getElementById("historyList");
  const conversationTitle = document.getElementById("conversationTitle");
  const conversationSubtitle = document.getElementById("conversationSubtitle");
  const modeLabel = document.getElementById("modeLabel");
  const modeButtons = Array.from(document.querySelectorAll(".mode-btn"));

  let conversations = loadConversations();
  let currentConversationId = conversations[0]?.id || null;
  let selectedFiles = [];
  let currentMode = "chat";

  renderHistory();
  renderCurrentConversation();

  if (usageModal && conversations.length === 0) {
    usageModal.hidden = false;
  }

  if (closeUsageTip && usageModal) {
    closeUsageTip.onclick = function () {
      usageModal.hidden = true;
    };
  }

  if (openUsageTip && usageModal) {
    openUsageTip.onclick = function () {
      usageModal.hidden = false;
    };
  }

  if (newChatBtn) {
    newChatBtn.onclick = function () {
      currentConversationId = null;
      selectedFiles = [];
      reportInput.value = "";
      if (fileUpload) fileUpload.value = "";
      renderFileList();
      renderCurrentConversation();
      reportInput.focus();
    };
  }

  modeButtons.forEach(function (button) {
    button.onclick = function () {
      setMode(button.dataset.mode || "chat");
    };
  });

  if (fileUpload && fileList) {
    fileUpload.onchange = function () {
      const files = Array.from(fileUpload.files);

      files.forEach(function (file) {
        const exists = selectedFiles.some(function (item) {
          return item.name === file.name && item.size === file.size;
        });

        if (!exists) {
          selectedFiles.push({
            name: file.name,
            size: file.size,
            type: file.type || "未知类型"
          });
        }
      });

      renderFileList();
      fileUpload.value = "";
    };
  }

  if (sendBtn && reportInput) {
    sendBtn.onclick = sendCurrentMessage;

    reportInput.addEventListener("keydown", function (event) {
      if (event.key === "Enter" && !event.shiftKey) {
        event.preventDefault();
        sendCurrentMessage();
      }
    });
  }

  async function sendCurrentMessage() {
    const text = reportInput.value.trim();

    if (!text && selectedFiles.length === 0) {
      alert("请先输入问题、PPT要求或上传附件。");
      return;
    }

    if (!text) {
      alert("当前版本暂未解析附件内容，请先输入文字说明。");
      return;
    }

    const conversation = ensureConversation(text);
    const files = selectedFiles.map(function (file) {
      return { name: file.name, size: file.size, type: file.type };
    });
    const userMessage = {
      role: "user",
      content: buildUserContent(text, files),
      rawText: text,
      files,
      mode: currentMode,
      createdAt: new Date().toISOString()
    };

    conversation.mode = currentMode;
    conversation.messages.push(userMessage);
    conversation.updatedAt = new Date().toISOString();

    reportInput.value = "";
    selectedFiles = [];
    renderFileList();
    renderCurrentConversation();
    persistAndRenderHistory();

    const loadingMessage = {
      role: "assistant",
      content: currentMode === "ppt" ? "正在生成PPT文件，请稍候..." : "正在思考...",
      pending: true,
      createdAt: new Date().toISOString()
    };
    conversation.messages.push(loadingMessage);
    renderCurrentConversation();

    setSending(true);

    try {
      const responseData =
        currentMode === "ppt"
          ? await requestPpt(conversation)
          : await requestChat(conversation);

      loadingMessage.pending = false;
      loadingMessage.content = responseData.content;
      loadingMessage.ppt = responseData.ppt || null;
      loadingMessage.createdAt = new Date().toISOString();
      conversation.updatedAt = new Date().toISOString();
      persistAndRenderHistory();
      renderCurrentConversation();
    } catch (error) {
      loadingMessage.pending = false;
      loadingMessage.error = true;
      loadingMessage.content = error.message || "请求失败，请稍后重试。";
      conversation.updatedAt = new Date().toISOString();
      persistAndRenderHistory();
      renderCurrentConversation();
    } finally {
      setSending(false);
    }
  }

  async function requestChat(conversation) {
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        mode: "chat",
        messages: toApiMessages(conversation.messages)
      })
    });

    const data = await response.json();
    if (!response.ok || !data.success) {
      throw new Error(data.message || "通用问答请求失败。");
    }

    return {
      content: data.result || "模型未返回内容。"
    };
  }

  async function requestPpt(conversation) {
    const response = await fetch("/api/generate-ppt", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        title: conversation.title,
        messages: toApiMessages(conversation.messages)
      })
    });

    const data = await response.json();
    if (!response.ok || !data.success) {
      throw new Error(data.message || "PPT生成失败。");
    }

    const slideSummary = Array.isArray(data.slides)
      ? data.slides.map(function (slide, index) {
          return `${index + 1}. ${slide.title}`;
        })
      : [];

    return {
      content: [data.message || "PPT已生成。", "", "页面结构：", ...slideSummary].join("\n"),
      ppt: {
        fileName: data.fileName,
        base64: data.pptxBase64
      }
    };
  }

  function setSending(isSending) {
    sendBtn.disabled = isSending;
    sendBtn.textContent = isSending ? "…" : "→";
  }

  function setMode(mode) {
    currentMode = mode === "ppt" ? "ppt" : "chat";

    modeButtons.forEach(function (button) {
      button.classList.toggle("active", button.dataset.mode === currentMode);
    });

    if (modeLabel) {
      modeLabel.textContent = currentMode === "ppt" ? "PPT制作" : "通用问答";
    }

    if (reportInput) {
      reportInput.placeholder =
        currentMode === "ppt"
          ? "输入PPT主题、受众、页数、材料或汇报重点"
          : "输入你的问题，或让智能体整理材料";
    }
  }

  function ensureConversation(firstText) {
    let conversation = getCurrentConversation();
    if (conversation) {
      return conversation;
    }

    conversation = {
      id: createId(),
      title: makeTitle(firstText),
      mode: currentMode,
      messages: [],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    conversations.unshift(conversation);
    currentConversationId = conversation.id;
    return conversation;
  }

  function renderCurrentConversation() {
    const conversation = getCurrentConversation();
    const hasMessages = conversation && conversation.messages.length > 0;

    chatArea.innerHTML = "";
    chatArea.classList.toggle("active", Boolean(hasMessages));
    mainContent.classList.toggle("chat-mode", Boolean(hasMessages));

    if (!conversation) {
      conversationTitle.textContent = "AI 汇报助手";
      conversationSubtitle.textContent = "可进行通用问答，也可以根据材料生成可下载的PPT文件。";
      setMode(currentMode);
      return;
    }

    conversationTitle.textContent = conversation.title || "未命名对话";
    conversationSubtitle.textContent = formatConversationMeta(conversation);
    setMode(conversation.mode || currentMode);

    conversation.messages.forEach(function (message) {
      appendMessageElement(message);
    });

    scrollChatToBottom();
  }

  function appendMessageElement(message) {
    const row = document.createElement("div");
    row.className = `message-row ${message.role === "user" ? "user" : "assistant"}`;

    const bubble = document.createElement("div");
    bubble.className = "message-bubble";
    if (message.pending) bubble.classList.add("pending");
    if (message.error) bubble.classList.add("error-text");

    const content = document.createElement("div");
    content.textContent = message.rawText || message.content;
    bubble.appendChild(content);

    if (message.files && message.files.length > 0) {
      const filesBox = document.createElement("div");
      filesBox.className = "message-files";

      message.files.forEach(function (file) {
        const chip = document.createElement("span");
        chip.className = "message-file-chip";
        chip.textContent = file.name;
        filesBox.appendChild(chip);
      });

      bubble.appendChild(filesBox);
    }

    if (message.ppt && message.ppt.base64) {
      const downloadLink = document.createElement("a");
      downloadLink.className = "ppt-download";
      downloadLink.href = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${message.ppt.base64}`;
      downloadLink.download = message.ppt.fileName || "智能体生成PPT.pptx";
      downloadLink.textContent = "下载PPT";
      bubble.appendChild(downloadLink);
    }

    row.appendChild(bubble);
    chatArea.appendChild(row);
  }

  function renderHistory() {
    historyList.innerHTML = "";

    const visibleConversations = conversations.filter(function (conversation) {
      return conversation.messages.length > 0;
    });

    if (visibleConversations.length === 0) {
      const empty = document.createElement("div");
      empty.className = "history-empty";
      empty.textContent = "暂无对话记录";
      historyList.appendChild(empty);
      return;
    }

    visibleConversations.forEach(function (conversation) {
      const item = document.createElement("button");
      item.className = "history-item";
      item.classList.toggle("active", conversation.id === currentConversationId);
      item.type = "button";
      item.onclick = function () {
        currentConversationId = conversation.id;
        currentMode = conversation.mode || "chat";
        selectedFiles = [];
        renderFileList();
        renderCurrentConversation();
        renderHistory();
      };

      const title = document.createElement("span");
      title.className = "history-title";
      title.textContent = conversation.title || "未命名对话";

      const meta = document.createElement("span");
      meta.className = "history-meta";
      meta.textContent = `${conversation.mode === "ppt" ? "PPT" : "问答"} · ${formatTime(conversation.updatedAt)}`;

      item.appendChild(title);
      item.appendChild(meta);
      historyList.appendChild(item);
    });
  }

  function renderFileList() {
    fileList.innerHTML = "";

    selectedFiles.forEach(function (file, index) {
      const fileTag = document.createElement("div");
      fileTag.className = "file-tag";

      const fileName = document.createElement("span");
      fileName.textContent = file.name;

      const removeBtn = document.createElement("button");
      removeBtn.className = "file-remove-btn";
      removeBtn.type = "button";
      removeBtn.textContent = "×";
      removeBtn.title = "删除附件";

      removeBtn.onclick = function () {
        selectedFiles.splice(index, 1);
        renderFileList();
      };

      fileTag.appendChild(fileName);
      fileTag.appendChild(removeBtn);
      fileList.appendChild(fileTag);
    });
  }

  function persistAndRenderHistory() {
    saveConversations();
    renderHistory();
  }

  function loadConversations() {
    try {
      const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  function saveConversations() {
    const stored = conversations
      .filter(function (conversation) {
        return conversation.messages.length > 0;
      })
      .slice(0, 50);

    localStorage.setItem(STORAGE_KEY, JSON.stringify(stored));
    conversations = stored;
  }

  function getCurrentConversation() {
    return conversations.find(function (conversation) {
      return conversation.id === currentConversationId;
    });
  }

  function toApiMessages(messages) {
    return messages
      .filter(function (message) {
        return !message.pending && !message.error && message.content;
      })
      .map(function (message) {
        return {
          role: message.role,
          content: message.content
        };
      });
  }

  function buildUserContent(text, files) {
    if (!files || files.length === 0) {
      return text;
    }

    const fileText = files
      .map(function (file) {
        return `- ${file.name}`;
      })
      .join("\n");

    return `${text}\n\n用户上传了以下附件（当前仅提供文件名，未解析正文）：\n${fileText}`;
  }

  function makeTitle(text) {
    return text
      .replace(/\s+/g, " ")
      .replace(/[。！？!?].*$/, "")
      .trim()
      .slice(0, 28) || "新对话";
  }

  function formatConversationMeta(conversation) {
    const modeText = conversation.mode === "ppt" ? "PPT制作" : "通用问答";
    return `${modeText} · ${conversation.messages.length} 条消息 · ${formatTime(conversation.updatedAt)}`;
  }

  function formatTime(value) {
    if (!value) return "刚刚";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "刚刚";

    return date.toLocaleString("zh-CN", {
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit"
    });
  }

  function createId() {
    return `conv_${Date.now()}_${Math.random().toString(16).slice(2)}`;
  }

  function scrollChatToBottom() {
    chatArea.scrollTop = chatArea.scrollHeight;
  }
});
