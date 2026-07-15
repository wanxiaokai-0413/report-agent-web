const express = require("express");
const path = require("path");
const fs = require("fs");
const zlib = require("zlib");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 3000;
const DEEPSEEK_MODEL = process.env.DEEPSEEK_MODEL || "deepseek-chat";
const DEEPSEEK_BASE_URL = process.env.DEEPSEEK_BASE_URL || "https://api.deepseek.com";
const DEEPSEEK_TIMEOUT_MS = Number(process.env.DEEPSEEK_TIMEOUT_MS || 30000);
const DEEPSEEK_MAX_RETRIES = Number(process.env.DEEPSEEK_MAX_RETRIES || 2);
const PPT_TEMPLATE_PATH = process.env.PPT_TEMPLATE_PATH || path.join(__dirname, "templates", "联想内部使用模板.pptx");
const LENOVO_WHITE_PROMPT_PATH = path.join(__dirname, "prompts", "lenovo-white-layout-prompt.txt");
const PUBLIC_DIR = path.join(__dirname, "public");
const LENOVO_WHITE_LAYOUTS = {
  "Title Slide_White": 1,
  "Section Header_White": 3,
  "仅标题": 5,
  "Title with Subtitle Only": 7,
  "标题和内容": 9,
  "Title with Subtitle Content": 11,
  "Two Column Slide": 13,
  "Three Column Slide": 15,
  "Title w/Image": 17,
  "Photo + Statement": 19,
  "Big Idea": 5,
  "Content w/ Product": 22,
  "Chart Slide": 24,
  "Blank Slide": 26,
  "Closing Slide": 28
};
const DEEPSEEK_MODEL_FALLBACKS = (process.env.DEEPSEEK_MODEL_FALLBACKS || "deepseek-chat")
  .split(",")
  .map((model) => model.trim())
  .filter(Boolean);

app.use(express.json({ limit: "20mb" }));
app.use(express.static(PUBLIC_DIR, { dotfiles: "ignore" }));

app.get("/api/ping", (req, res) => {
  res.json({
    success: true,
    message: "后端服务正常运行"
  });
});

app.post("/api/chat", async (req, res) => {
  try {
    const { messages = [], mode = "chat" } = req.body;

    if (!Array.isArray(messages) || messages.length === 0) {
      return res.status(400).json({
        success: false,
        message: "请先输入对话内容。"
      });
    }

    let result;
    try {
      result = await callDeepSeek({
        messages: [
          {
            role: "system",
            content:
              mode === "ppt"
                ? "你是一个擅长制作企业汇报PPT的智能体。先理解用户目标，再给出清晰、可执行、适合做成幻灯片的内容建议。"
                : "你是汇报智能体，一个接入DeepSeek大模型的通用问答、材料整理和PPT制作助手。用户问你是谁或你能做什么时，说明你的产品身份是汇报智能体，而不是直接自称DeepSeek。回答应准确、清晰、直接。遇到不确定的信息要说明不确定性，不要编造事实。"
          },
          ...normalizeMessages(messages)
        ],
        temperature: mode === "ppt" ? 0.35 : 0.7,
        max_tokens: 2400
      });
    } catch (error) {
      const fallback = buildLocalChatFallback(messages, mode);
      if (!fallback) {
        throw error;
      }
      console.warn("DeepSeek unavailable, using local chat fallback:", error.message);
      result = fallback;
    }

    res.json({
      success: true,
      result
    });
  } catch (error) {
    sendServerError(res, error);
  }
});

app.post("/api/generate-report", async (req, res) => {
  try {
    const { reportText } = req.body;

    if (!reportText || !reportText.trim()) {
      return res.status(400).json({
        success: false,
        message: "请输入需要整理的工作内容。"
      });
    }

    let result;
    try {
      result = await callDeepSeek({
        messages: [
          {
            role: "system",
            content:
              "你是专业的企业工作汇报整理助手。将零散工作内容整理成正式、清晰、适合后续制作PPT的结构化汇报草稿。不要编造用户没有提供的事实、数据或项目名称。"
          },
          {
            role: "user",
            content: `请把以下工作内容整理成汇报草稿，包含：汇报标题、总体概述、重点工作进展、阶段成果、存在问题、下阶段计划。\n\n${reportText}`
          }
        ],
        temperature: 0.4,
        max_tokens: 2200
      });
    } catch (error) {
      console.warn("DeepSeek unavailable, using local report fallback:", error.message);
      result = buildLocalReport(reportText);
    }

    res.json({
      success: true,
      result
    });
  } catch (error) {
    sendServerError(res, error);
  }
});

app.post("/api/generate-ppt", async (req, res) => {
  try {
    const { messages = [], title } = req.body;

    if (!Array.isArray(messages) || messages.length === 0) {
      return res.status(400).json({
        success: false,
        message: "请先提供PPT主题或材料。"
      });
    }

    let deck;
    try {
      const pptJsonText = await callDeepSeek({
        messages: [
          {
            role: "system",
            content: getPptPlanningPrompt()
          },
          ...normalizeMessages(messages)
        ],
        temperature: 0.35,
        max_tokens: 4500
      });

      deck = parseDeckJson(pptJsonText, title);
    } catch (error) {
      console.warn("DeepSeek unavailable, using local PPT fallback:", error.message);
      deck = buildLocalPptDeck(messages, title);
    }
    const pptxBuffer = fs.existsSync(PPT_TEMPLATE_PATH)
      ? buildLenovoPptxFromTemplate(deck, PPT_TEMPLATE_PATH)
      : buildPptx(deck);
    const safeTitle = sanitizeFileName(deck.title || title || "智能体生成PPT");

    res.json({
      success: true,
      title: deck.title,
      slides: deck.slides,
      fileName: `${safeTitle}.pptx`,
      pptxBase64: pptxBuffer.toString("base64"),
      message: `已生成 ${deck.slides.length} 页PPT，可点击下载。`
    });
  } catch (error) {
    sendServerError(res, error);
  }
});

async function callDeepSeek({ messages, temperature, max_tokens }) {
  if (!process.env.DEEPSEEK_API_KEY) {
    const error = new Error("后端未配置 DeepSeek API Key，请检查 .env 文件。");
    error.statusCode = 500;
    throw error;
  }

  const models = uniqueValues([DEEPSEEK_MODEL, ...DEEPSEEK_MODEL_FALLBACKS]);
  let lastError;

  for (const model of models) {
    for (let attempt = 0; attempt <= DEEPSEEK_MAX_RETRIES; attempt += 1) {
      try {
        return await requestDeepSeekOnce({
          model,
          messages,
          temperature,
          max_tokens
        });
      } catch (error) {
        lastError = error;

        if (!isRetryableDeepSeekError(error) || attempt === DEEPSEEK_MAX_RETRIES) {
          break;
        }

        await sleep(500 * (attempt + 1));
      }
    }
  }

  throw lastError;
}

async function requestDeepSeekOnce({ model, messages, temperature, max_tokens }) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), DEEPSEEK_TIMEOUT_MS);

  try {
    const response = await fetch(`${DEEPSEEK_BASE_URL}/chat/completions`, {
      method: "POST",
      signal: controller.signal,
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${process.env.DEEPSEEK_API_KEY}`
      },
      body: JSON.stringify({
        model,
        messages,
        temperature,
        max_tokens
      })
    });

    const data = await response.json().catch(() => ({}));

    if (!response.ok) {
      const error = new Error(`DeepSeek API 调用失败：${JSON.stringify(data)}`);
      error.statusCode = response.status;
      error.detail = data;
      throw error;
    }

    return data.choices?.[0]?.message?.content?.trim() || "";
  } catch (error) {
    if (error.name === "AbortError") {
      const timeoutError = new Error(`DeepSeek 请求超过 ${DEEPSEEK_TIMEOUT_MS}ms 未响应。`);
      timeoutError.code = "ETIMEDOUT";
      throw timeoutError;
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

function isRetryableDeepSeekError(error) {
  if (!error) {
    return false;
  }

  if (["ETIMEDOUT", "ECONNRESET", "ECONNREFUSED", "EAI_AGAIN", "ENOTFOUND", "EACCES"].includes(error.code)) {
    return true;
  }

  if (error.message === "fetch failed") {
    return true;
  }

  if (typeof error.statusCode === "number") {
    return error.statusCode === 408 || error.statusCode === 429 || error.statusCode >= 500;
  }

  return false;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function uniqueValues(values) {
  return [...new Set(values.filter(Boolean))];
}

function normalizeMessages(messages) {
  return messages
    .filter((message) => message && typeof message.content === "string")
    .map((message) => ({
      role: message.role === "assistant" ? "assistant" : "user",
      content: message.content.slice(0, 12000)
    }))
    .slice(-12);
}

function buildLocalChatFallback(messages, mode) {
  const lastUserText = getLastUserText(messages);
  const normalized = lastUserText.toLowerCase().replace(/\s+/g, "");

  if (!lastUserText.trim()) {
    return null;
  }

  if (
    /你是谁|你叫什么|介绍一下你|你能做什么|whoareyou|whatareyou/.test(normalized)
  ) {
    return [
      "我是汇报智能体，一个面向工作问答、材料整理和PPT初稿生成的AI助手。",
      "我可以进行通用问答，也可以把工作内容整理成汇报结构，并生成可下载的PPT文件。"
    ].join("\n");
  }

  if (/你好|您好|hello|hi|嗨/.test(normalized)) {
    return "你好，我是汇报智能体。你可以直接向我提问，也可以切换到“生成PPT”模式，让我根据材料制作PPT初稿。";
  }

  if (mode === "ppt" || /ppt|幻灯片|演示文稿|汇报材料/.test(normalized)) {
    return [
      "我可以帮你制作PPT初稿。请提供主题、受众、页数、材料要点和希望的风格。",
      "如果你已经有工作内容，可以直接粘贴，我会整理成标题页、背景目标、进展成果、问题风险、后续计划等页面结构。"
    ].join("\n");
  }

  if (/周报|月报|季报|汇报|总结|工作内容/.test(normalized)) {
    return buildLocalReport(lastUserText);
  }

  return [
    "当前大模型服务暂时不可用，我无法可靠完成开放式问答。",
    "但我仍可以处理基础身份说明、使用指引、汇报草稿整理和PPT初稿生成。请稍后重试，或先提供工作材料让我整理。"
  ].join("\n");
}

function buildLocalReport(reportText) {
  const points = extractContentPoints(reportText);

  return [
    "# 工作汇报草稿",
    "",
    "## 总体概述",
    points.length > 0
      ? `本阶段围绕 ${points.slice(0, 2).join("、")} 等内容推进，整体工作按材料记录持续开展。`
      : "本阶段工作内容已收到，可在补充更多细节后进一步完善。",
    "",
    "## 重点工作进展",
    ...formatBulletLines(points.slice(0, 5)),
    "",
    "## 阶段成果",
    points.length > 0 ? "- 已完成或推进了材料中提到的核心事项。" : "- 暂无明确成果信息，请补充具体产出。",
    "",
    "## 存在问题",
    "- 请结合实际补充当前风险、阻塞点和资源需求。",
    "",
    "## 下阶段计划",
    "- 明确下一阶段目标、责任人和时间节点。",
    "- 对关键风险进行跟踪，并及时同步进展。"
  ].join("\n");
}

function buildLocalPptDeck(messages, fallbackTitle) {
  const text = getLastUserText(messages);
  const parsed = parseReportForTemplate(text);
  const title = parsed.title || fallbackTitle || makeDeckTitle(text);
  const capacities = { review: 4, projects: 3, problems: 3, plans: 3 };
  const slides = [];

  for (const category of ["review", "projects", "problems", "plans"]) {
    const items = parsed.categories[category];
    const selected = items.slice(0, capacities[category]);
    for (const item of selected) {
      slides.push({
        title: compactTitle(item.title, 14),
        bullets: item.bullets.slice(0, 5).map((bullet) => compactBullet(bullet, 46)),
        notes: ""
      });
    }
    while (selected.length < capacities[category]) {
      slides.push({
        title: categoryDefaultTitle(category, selected.length),
        bullets: ["暂无更多可填充内容"],
        notes: ""
      });
      selected.push({});
    }
  }

  return { title, slides };
}

function parseReportForTemplate(text) {
  const lines = String(text || "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const categories = { review: [], projects: [], problems: [], plans: [] };
  let category = "review";
  let current = null;
  let title = "";

  const flush = () => {
    if (!current) return;
    current.bullets = current.bullets
      .flatMap(splitReportSentence)
      .map((item) => item.replace(/^(?:[-*•]\s*|\d+[.、)）]\s*)/, "").trim())
      .filter((item) => item.length > 1);
    if (current.title && current.bullets.length) categories[category].push(current);
    current = null;
  };

  for (const rawLine of lines) {
    const line = rawLine.replace(/^#+\s*/, "").trim();
    if (!title && /工作总结|工作汇报|汇报$/.test(line)) title = line.replace(/^[-—\s]+|[-—\s]+$/g, "");

    const nextCategory = detectReportCategory(line);
    if (nextCategory && isMajorReportHeading(rawLine, line)) {
      flush();
      category = nextCategory;
      continue;
    }

    if (isReportSubheading(rawLine, line)) {
      flush();
      current = { title: cleanReportHeading(line), bullets: [] };
      continue;
    }

    if (!current) current = { title: categoryDefaultTitle(category, categories[category].length), bullets: [] };
    if (!/^汇报(部门|周期|人|日期)[:：]/.test(line) && line !== title) current.bullets.push(line);
  }
  flush();
  return { title, categories };
}

function detectReportCategory(line) {
  if (/不足|问题|风险|改进方案/.test(line)) return "problems";
  if (/下一阶段|下阶段|工作计划|时间安排|预期成果|协调与支持/.test(line)) return "plans";
  if (/重点项目|项目取得进展|阶段性成果|经营指标/.test(line)) return "projects";
  if (/季度工作|工作复盘|本季度|工作概述|重点目标|完成情况|产品运营|客户运营|内容与活动|团队协作/.test(line)) return "review";
  return "";
}

function isMajorReportHeading(rawLine, line) {
  return /^#\s+/.test(rawLine) || /^[一二三四五六七八九十]+[、.．]/.test(line);
}

function isReportSubheading(rawLine, line) {
  return /^#{2,}\s+/.test(rawLine) || /^\d+[、.．]\s*/.test(line);
}

function cleanReportHeading(line) {
  return line.replace(/^[一二三四五六七八九十\d]+[、.．]\s*/, "").trim().slice(0, 28) || "工作内容";
}

function splitReportSentence(text) {
  const value = String(text || "").trim();
  return value.length <= 64 ? [value] : value.split(/[。；;]/).map((item) => item.trim()).filter(Boolean);
}

function categoryDefaultTitle(category, index) {
  const titles = {
    review: ["本季度工作概述", "重点目标与任务", "运营工作进展", "团队协作与管理"],
    projects: ["重点项目进展", "关键经营指标", "阶段性成果"],
    problems: ["主要问题与不足", "原因分析", "改进措施与风险"],
    plans: ["下一阶段总体目标", "重点工作计划", "时间安排与预期成果"]
  };
  return titles[category]?.[index] || "工作内容";
}

function getLastUserText(messages) {
  const userMessages = Array.isArray(messages) ? messages.filter((message) => message?.role !== "assistant") : [];
  return userMessages.at(-1)?.content || "";
}

function extractContentPoints(text) {
  const lines = String(text || "")
    .split(/[\r\n。；;]+/)
    .map((line) => line.replace(/^[-*#\d.\s]+/, "").trim())
    .filter(Boolean)
    .filter((line) => line.length > 2);

  if (lines.length > 0) {
    return lines.slice(0, 8);
  }

  const compact = String(text || "").trim();
  return compact ? [compact.slice(0, 80)] : ["请补充具体材料要点"];
}

function formatBulletLines(points) {
  if (points.length === 0) {
    return ["- 请补充具体工作事项。"];
  }

  return points.map((point) => `- ${point}`);
}

function makeDeckTitle(text) {
  const firstPoint = extractContentPoints(text)[0] || "智能体生成PPT";
  return firstPoint.replace(/[，,：:；;。.!！?？].*$/, "").slice(0, 24) || "智能体生成PPT";
}

function getPptPlanningPrompt() {
  if (fs.existsSync(LENOVO_WHITE_PROMPT_PATH)) {
    return fs.readFileSync(LENOVO_WHITE_PROMPT_PATH, "utf8");
  }

  return [
    "你是专业的PPT策划与内容生成智能体。",
    "请根据用户材料生成一份企业汇报PPT结构。",
    "只输出合法JSON，不要输出Markdown代码块，不要输出解释。",
    "JSON格式：",
    '{"title":"PPT标题","slides":[{"title":"页标题","bullets":["要点1","要点2","要点3"],"notes":"演讲备注"}]}',
    "要求：6到10页；每页2到5个要点；内容简洁、正式；不要编造材料中没有的具体事实或数据。"
  ].join("\n");
}

function parseDeckJson(text, fallbackTitle) {
  const cleaned = String(text || "")
    .replace(/^```json/i, "")
    .replace(/^```/, "")
    .replace(/```$/, "")
    .trim();

  const start = cleaned.indexOf("{");
  const end = cleaned.lastIndexOf("}");
  const jsonText = start >= 0 && end > start ? cleaned.slice(start, end + 1) : cleaned;

  let parsed;
  try {
    parsed = JSON.parse(jsonText);
  } catch {
    parsed = fallbackDeckFromText(cleaned, fallbackTitle);
  }

  const title = String(parsed.title || parsed.deck_title || fallbackTitle || "智能体生成PPT").trim();
  const slides = Array.isArray(parsed.slides) ? parsed.slides : [];
  const normalizedSlides = slides
    .map((slide, index) => normalizeStructuredSlide(slide, index))
    .filter((slide) => slide.title || slide.bullets.length > 0)
    .slice(0, 12);

  if (normalizedSlides.length === 0) {
    normalizedSlides.push({
      title,
      bullets: ["请补充更完整的材料后重新生成。"],
      notes: ""
    });
  }

  return {
    title,
    subtitle: String(parsed.deck_subtitle || parsed.subtitle || "").trim(),
    slides: normalizedSlides
  };
}

function normalizeStructuredSlide(slide, index) {
  const title = String(
    slide.title ||
      slide.statement ||
      slide.idea ||
      slide.closing_text ||
      slide.chart_title ||
      `第 ${index + 1} 页`
  ).trim();
  const bullets = [];

  appendBullets(bullets, slide.subtitle ? [slide.subtitle] : []);
  appendBullets(bullets, slide.bullets);

  if (slide.left_heading || Array.isArray(slide.left_bullets)) {
    appendBullets(bullets, [slide.left_heading]);
    appendBullets(bullets, slide.left_bullets);
  }

  if (slide.right_heading || Array.isArray(slide.right_bullets)) {
    appendBullets(bullets, [slide.right_heading]);
    appendBullets(bullets, slide.right_bullets);
  }

  if (Array.isArray(slide.columns)) {
    for (const column of slide.columns) {
      appendBullets(bullets, [column?.heading]);
      appendBullets(bullets, column?.bullets);
    }
  }

  appendBullets(bullets, slide.insight ? [slide.insight] : []);
  appendBullets(bullets, slide.image_description ? [`图片建议：${slide.image_description}`] : []);
  appendBullets(bullets, slide.custom_content_plan ? [`自定义内容：${slide.custom_content_plan}`] : []);

  return {
    title,
    subtitle: String(slide.subtitle || "").trim(),
    layout: String(slide.layout || "").trim(),
    purpose: String(slide.purpose || "").trim(),
    statement: String(slide.statement || "").trim(),
    idea: String(slide.idea || "").trim(),
    closing_text: String(slide.closing_text || "").trim(),
    left_heading: String(slide.left_heading || "").trim(),
    left_bullets: normalizeStringArray(slide.left_bullets).slice(0, 5),
    right_heading: String(slide.right_heading || "").trim(),
    right_bullets: normalizeStringArray(slide.right_bullets).slice(0, 5),
    columns: Array.isArray(slide.columns)
      ? slide.columns.slice(0, 3).map((column) => ({
          heading: String(column?.heading || "").trim(),
          bullets: normalizeStringArray(column?.bullets).slice(0, 4)
        }))
      : [],
    chart_type: String(slide.chart_type || "").trim(),
    chart_title: String(slide.chart_title || "").trim(),
    chart_data: slide.chart_data && typeof slide.chart_data === "object" ? slide.chart_data : null,
    insight: String(slide.insight || "").trim(),
    image_description: String(slide.image_description || "").trim(),
    custom_content_plan: String(slide.custom_content_plan || "").trim(),
    bullets: bullets.slice(0, 6),
    notes: String(slide.notes || slide.speaker_notes || "").trim()
  };
}

function appendBullets(target, values) {
  if (!Array.isArray(values)) return;
  for (const value of values) {
    const text = String(value || "").trim();
    if (text) target.push(text);
  }
}

function normalizeStringArray(values) {
  if (!Array.isArray(values)) return [];
  return values.map((value) => String(value || "").trim()).filter(Boolean);
}

function fallbackDeckFromText(text, fallbackTitle) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((line) => line.replace(/^[-*#\d.\s]+/, "").trim())
    .filter(Boolean);

  const title = fallbackTitle || lines[0] || "智能体生成PPT";
  const bullets = lines.slice(1, 16);
  const slides = [];

  for (let i = 0; i < Math.max(1, Math.ceil(bullets.length / 4)); i += 1) {
    slides.push({
      title: i === 0 ? title : `内容要点 ${i + 1}`,
      bullets: bullets.slice(i * 4, i * 4 + 4),
      notes: ""
    });
  }

  return { title, slides };
}

function buildLenovoPptxFromTemplate(deck, templatePath) {
  const templateEntries = readZipEntries(fs.readFileSync(templatePath));
  const preparedDeck = prepareLenovoDeck(deck);
  const slideCount = preparedDeck.slides.length;
  const files = templateEntries.filter(([name]) => !shouldReplaceForGeneratedDeck(name));

  files.push(["[Content_Types].xml", contentTypesFromTemplate(templateEntries, slideCount)]);
  files.push(["_rels/.rels", rootRelsXml()]);
  files.push(["docProps/core.xml", coreXml(preparedDeck.title)]);
  files.push(["docProps/app.xml", appXml(slideCount)]);
  files.push(["ppt/presentation.xml", lenovoPresentationXml(slideCount)]);
  files.push(["ppt/_rels/presentation.xml.rels", presentationRelsXml(slideCount)]);

  preparedDeck.slides.forEach((slide, index) => {
    const slideNumber = index + 1;
    const layout = resolveLenovoLayout(slide, index, preparedDeck.slides.length);
    const layoutNumber = LENOVO_WHITE_LAYOUTS[layout] || LENOVO_WHITE_LAYOUTS["标题和内容"];
    files.push([`ppt/slides/slide${slideNumber}.xml`, lenovoSlideXml(slide, slideNumber, layout, preparedDeck)]);
    files.push([`ppt/slides/_rels/slide${slideNumber}.xml.rels`, lenovoSlideRelsXml(layoutNumber)]);
  });

  return zipStore(files);
}

function shouldReplaceForGeneratedDeck(name) {
  return (
    name === "[Content_Types].xml" ||
    name === "_rels/.rels" ||
    name === "docProps/core.xml" ||
    name === "docProps/app.xml" ||
    name === "ppt/presentation.xml" ||
    name === "ppt/_rels/presentation.xml.rels" ||
    /^ppt\/slides\//.test(name) ||
    /^ppt\/notesSlides\//.test(name)
  );
}

function contentTypesFromTemplate(templateEntries, slideCount) {
  const source = templateEntries.find(([name]) => name === "[Content_Types].xml")?.[1]?.toString("utf8");
  const base = source || contentTypesXml(0);
  const slideOverrides = Array.from({ length: slideCount }, (_, index) => {
    const slideNumber = index + 1;
    return `<Override PartName="/ppt/slides/slide${slideNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  }).join("");

  return base
    .replace(/\s*<Override PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, "")
    .replace(/\s*<Override PartName="\/ppt\/notesSlides\/notesSlide\d+\.xml"[^>]*\/>/g, "")
    .replace("</Types>", `${slideOverrides}</Types>`);
}

function prepareLenovoDeck(deck) {
  const sourceSlides = Array.isArray(deck.slides) ? deck.slides.filter(Boolean) : [];
  const title = String(deck.title || "智能体生成PPT").trim();
  const subtitle = String(deck.subtitle || "").trim();
  const slides = sourceSlides.map((slide) => ({ ...slide }));

  if (slides.length === 0 || resolveLenovoLayout(slides[0], 0, slides.length) !== "Title Slide_White") {
    slides.unshift({
      layout: "Title Slide_White",
      title,
      subtitle: subtitle || "汇报材料",
      bullets: []
    });
  }

  if (resolveLenovoLayout(slides[slides.length - 1], slides.length - 1, slides.length) !== "Closing Slide") {
    slides.push({
      layout: "Closing Slide",
      title: "Thanks",
      closing_text: "Thanks",
      bullets: []
    });
  }

  return { title, subtitle, slides: slides.slice(0, 18) };
}

function resolveLenovoLayout(slide, index, total) {
  const requested = String(slide?.layout || "").trim();
  if (Object.prototype.hasOwnProperty.call(LENOVO_WHITE_LAYOUTS, requested)) return requested;
  if (index === 0) return "Title Slide_White";
  if (index === total - 1) return "Closing Slide";
  if (slide?.chart_data || slide?.chart_type) return "Chart Slide";
  if (Array.isArray(slide?.columns) && slide.columns.length >= 3) return "Three Column Slide";
  if (slide?.left_heading || slide?.right_heading || slide?.left_bullets || slide?.right_bullets) return "Two Column Slide";
  if (slide?.idea) return "Big Idea";
  if (slide?.statement && slide?.image_description) return "Photo + Statement";
  if (slide?.image_description && String(slide?.purpose || "").includes("product")) return "Content w/ Product";
  if (slide?.image_description) return "Title w/Image";
  if (slide?.subtitle) return "Title with Subtitle Content";
  return "标题和内容";
}

function lenovoPresentationXml(slideCount) {
  const slideIds = Array.from({ length: slideCount }, (_, index) => {
    const slideNumber = index + 1;
    return `<p:sldId id="${255 + slideNumber}" r:id="rId${slideNumber + 1}"/>`;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${slideIds}</p:sldIdLst>
  <p:sldSz cx="12188825" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
}

function lenovoSlideRelsXml(layoutNumber) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${layoutNumber}.xml"/>
</Relationships>`;
}

function lenovoSlideXml(slide, slideNumber, layout, deck) {
  let id = 2;
  const nextId = () => id++;
  const shapes = [];
  const addPageNumber = () => {
    shapes.push(lenovoTextShape(nextId(), "Slide Number", String(slideNumber), 918.86, 504, 34.56, 12.24, 10, "1E0013", false, { align: "r" }));
  };

  switch (layout) {
    case "Title Slide_White":
      shapes.push(lenovoTextShape(nextId(), "Title", compactTitle(slide.title || deck.title, 34), 59.08, 120.29, 780.48, 257.46, 48, "1E0013", true));
      shapes.push(lenovoTextShape(nextId(), "Subtitle", slide.subtitle || deck.subtitle || "汇报材料", 59.88, 397.33, 780.48, 28.08, 18, "4E444E", false));
      break;
    case "Section Header_White":
      shapes.push(lenovoTextShape(nextId(), "Section Title", compactTitle(slide.title, 24), 59.08, 122.01, 780.48, 189.39, 44, "1E0013", true));
      shapes.push(lenovoTextShape(nextId(), "Section Subtitle", slide.subtitle || firstUsefulBullet(slide), 59.08, 340.84, 780.48, 25.75, 20, "4E444E", false));
      addPageNumber();
      break;
    case "仅标题":
      shapes.push(lenovoTextShape(nextId(), "Title", compactTitle(slide.title, 32), 59.87, 34.84, 840, 32.96, 32, "1E0013", true));
      addPageNumber();
      break;
    case "Title with Subtitle Only":
      shapes.push(lenovoTextShape(nextId(), "Title", compactTitle(slide.title, 32), 59.87, 34.84, 840, 32.96, 32, "1E0013", true));
      shapes.push(lenovoTextShape(nextId(), "Subtitle", slide.subtitle || firstUsefulBullet(slide), 59.87, 73.35, 840.24, 20.6, 20, "4E444E", false));
      addPageNumber();
      break;
    case "Title with Subtitle Content":
      shapes.push(lenovoTextShape(nextId(), "Title", compactTitle(slide.title, 32), 59.87, 34.84, 840.24, 32.96, 32, "1E0013", true));
      shapes.push(lenovoTextShape(nextId(), "Subtitle", slide.subtitle || slide.insight || "", 59.87, 73.35, 840.24, 20.6, 20, "4E444E", false));
      shapes.push(lenovoBulletShape(nextId(), slideBullets(slide, 5, 26), 59.87, 137.75, 840.24, 342.25, 20));
      addPageNumber();
      break;
    case "Two Column Slide":
      shapes.push(lenovoTextShape(nextId(), "Title", compactTitle(slide.title, 32), 59.87, 34.84, 840.24, 32.96, 32, "1E0013", true));
      shapes.push(lenovoColumnShape(nextId(), slide.left_heading || "现状", slide.left_bullets, 59.87, 102.96, 399.13, 377.04));
      shapes.push(lenovoColumnShape(nextId(), slide.right_heading || "计划", slide.right_bullets, 499.16, 102.96, 398.88, 377.04));
      addPageNumber();
      break;
    case "Three Column Slide":
      shapes.push(lenovoTextShape(nextId(), "Title", compactTitle(slide.title, 32), 59.87, 34.84, 840.24, 32.96, 32, "1E0013", true));
      [59.88, 344.67, 631.17].forEach((x, columnIndex) => {
        const column = slide.columns?.[columnIndex] || {};
        shapes.push(lenovoColumnShape(nextId(), column.heading || `模块 ${columnIndex + 1}`, column.bullets, x, 102.96, columnIndex === 2 ? 268.95 : 263.17, 377.04, 18));
      });
      addPageNumber();
      break;
    case "Title w/Image":
      shapes.push(lenovoTextShape(nextId(), "Title", compactTitle(slide.title, 22), 59.87, 34.84, 450.67, 32.96, 32, "1E0013", true));
      shapes.push(lenovoBulletShape(nextId(), slideBullets(slide, 5, 20), 59.87, 102.96, 450.67, 377.04, 20));
      shapes.push(lenovoImagePlaceholder(nextId(), slide.image_description, 539.88, 0, 419.87, 540));
      addPageNumber();
      break;
    case "Photo + Statement":
      shapes.push(lenovoTextShape(nextId(), "Statement", compactTitle(slide.statement || slide.title, 28), 59.87, 78.35, 269.85, 113.11, 30, "1E0013", true));
      shapes.push(lenovoBulletShape(nextId(), slideBullets(slide, 4, 18), 59.87, 221.73, 269.85, 258.27, 18));
      shapes.push(lenovoImagePlaceholder(nextId(), slide.image_description, 359.88, 0, 599.87, 540));
      addPageNumber();
      break;
    case "Big Idea":
      shapes.push(lenovoTextShape(nextId(), "Big Idea", compactTitle(slide.idea || slide.title, 28), 59.88, 138.41, 840, 263.19, 50, "4D144A", true));
      addPageNumber();
      break;
    case "Content w/ Product":
      shapes.push(lenovoImagePlaceholder(nextId(), slide.image_description, 0, 0, 479.87, 540));
      shapes.push(lenovoTextShape(nextId(), "Title", compactTitle(slide.title, 18), 539.87, 34.84, 360, 32.96, 32, "1E0013", true));
      shapes.push(lenovoBulletShape(nextId(), slideBullets(slide, 5, 18), 539.87, 102.96, 360, 377.04, 18));
      addPageNumber();
      break;
    case "Chart Slide":
      shapes.push(lenovoTextShape(nextId(), "Title", compactTitle(slide.title, 32), 59.87, 34.84, 840.24, 32.96, 32, "1E0013", true));
      shapes.push(...lenovoChartShapes(nextId, slide, 59.88, 120, 840, 360));
      addPageNumber();
      break;
    case "Blank Slide":
      shapes.push(lenovoTextShape(nextId(), "Custom Plan", slide.custom_content_plan || slide.title || "", 59.87, 103.06, 840.24, 376.94, 24, "4E444E", false));
      addPageNumber();
      break;
    case "Closing Slide":
      break;
    case "标题和内容":
    default:
      shapes.push(lenovoTextShape(nextId(), "Title", compactTitle(slide.title, 32), 59.87, 34.84, 840.24, 32.96, 32, "1E0013", true));
      shapes.push(lenovoBulletShape(nextId(), slideBullets(slide, 5, 28), 59.88, 103.06, 840.24, 376.94, 20));
      addPageNumber();
      break;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      ${shapes.join("\n")}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function slideBullets(slide, maxItems, maxLength) {
  return normalizeStringArray(slide.bullets).slice(0, maxItems).map((bullet) => compactBullet(bullet, maxLength));
}

function firstUsefulBullet(slide) {
  return slideBullets(slide, 1, 42)[0] || "";
}

function lenovoTextShape(id, name, text, x, y, w, h, sizePt, color, bold, options = {}) {
  const align = options.align || "l";
  const paragraphs = String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const paragraphXml = (paragraphs.length ? paragraphs : [""])
    .map((line) => `<a:p><a:pPr algn="${align}"/><a:r><a:rPr lang="zh-CN" sz="${Math.round(sizePt * 100)}"${bold ? ' b="1"' : ""}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Arial"/></a:rPr><a:t>${escapeXml(line)}</a:t></a:r></a:p>`)
    .join("");

  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${ptToEmu(x)}" y="${ptToEmu(y)}"/><a:ext cx="${ptToEmu(w)}" cy="${ptToEmu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>${paragraphXml}</p:txBody>
</p:sp>`;
}

function lenovoBulletShape(id, bullets, x, y, w, h, sizePt = 20) {
  const items = bullets.length ? bullets : ["请补充更完整的内容。"];
  const paragraphs = items
    .map((bullet) => `<a:p><a:pPr marL="228600" indent="-171450"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="zh-CN" sz="${Math.round(sizePt * 100)}"><a:solidFill><a:srgbClr val="1E0013"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Arial"/></a:rPr><a:t>${escapeXml(bullet)}</a:t></a:r></a:p>`)
    .join("");

  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="Bullets"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${ptToEmu(x)}" y="${ptToEmu(y)}"/><a:ext cx="${ptToEmu(w)}" cy="${ptToEmu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>${paragraphs}</p:txBody>
</p:sp>`;
}

function lenovoColumnShape(id, heading, bullets, x, y, w, h, sizePt = 20) {
  const items = normalizeStringArray(bullets).slice(0, 5).map((bullet) => compactBullet(bullet, 22));
  const headingXml = `<a:p><a:r><a:rPr lang="zh-CN" sz="${Math.round(sizePt * 100)}" b="1"><a:solidFill><a:srgbClr val="1E0013"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Arial"/></a:rPr><a:t>${escapeXml(heading)}</a:t></a:r></a:p>`;
  const bulletsXml = (items.length ? items : ["请补充内容"])
    .map((bullet) => `<a:p><a:pPr marL="228600" indent="-171450"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="zh-CN" sz="${Math.round((sizePt - 2) * 100)}"><a:solidFill><a:srgbClr val="1E0013"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Arial"/></a:rPr><a:t>${escapeXml(bullet)}</a:t></a:r></a:p>`)
    .join("");

  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="Column"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${ptToEmu(x)}" y="${ptToEmu(y)}"/><a:ext cx="${ptToEmu(w)}" cy="${ptToEmu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>${headingXml}${bulletsXml}</p:txBody>
</p:sp>`;
}

function lenovoImagePlaceholder(id, description, x, y, w, h) {
  const label = description ? `图片占位：${description}` : "图片占位";
  return `${lenovoRectShape(id, "Image Placeholder", x, y, w, h, "E6E2E4", "ABA8B1")}
${lenovoTextShape(id + 1000, "Image Description", compactBullet(label, 42), x + 24, y + h / 2 - 24, Math.max(60, w - 48), 48, 16, "4E444E", false, { align: "ctr" })}`;
}

function lenovoChartShapes(nextId, slide, x, y, w, h) {
  const shapes = [lenovoRectShape(nextId(), "Chart Background", x, y, w, h, "FFFFFF", "E6E2E4")];
  const chart = normalizeChartData(slide.chart_data);

  if (!chart.categories.length || !chart.values.length) {
    shapes.push(lenovoTextShape(nextId(), "Chart Empty", slide.insight || "请补充可视化数据。", x + 30, y + 120, w - 60, 80, 22, "4E444E", false, { align: "ctr" }));
    return shapes;
  }

  const max = Math.max(...chart.values, 1);
  const barAreaW = w - 180;
  const barH = Math.min(28, (h - 80) / chart.values.length - 8);
  chart.categories.slice(0, 8).forEach((category, index) => {
    const value = chart.values[index] || 0;
    const rowY = y + 38 + index * (barH + 12);
    const barW = Math.max(8, (value / max) * barAreaW);
    shapes.push(lenovoTextShape(nextId(), "Chart Label", compactBullet(category, 12), x + 24, rowY - 2, 120, barH + 6, 12, "1E0013", false));
    shapes.push(lenovoRectShape(nextId(), "Chart Bar", x + 150, rowY, barW, barH, index % 2 ? "4D144A" : "E1251B", "FFFFFF"));
    shapes.push(lenovoTextShape(nextId(), "Chart Value", String(value), x + 158 + barW, rowY - 2, 80, barH + 6, 12, "4E444E", false));
  });

  if (slide.insight) {
    shapes.push(lenovoTextShape(nextId(), "Chart Insight", compactBullet(slide.insight, 44), x + 24, y + h - 42, w - 48, 24, 14, "4E444E", false));
  }

  return shapes;
}

function normalizeChartData(chartData) {
  if (!chartData || typeof chartData !== "object") return { categories: [], values: [] };
  const categories = normalizeStringArray(chartData.categories).slice(0, 8);
  const rawSeries = Array.isArray(chartData.series) ? chartData.series[0] : null;
  const values = Array.isArray(rawSeries?.values)
    ? rawSeries.values
    : Array.isArray(chartData.values)
      ? chartData.values
      : [];
  return {
    categories,
    values: values.map((value) => Number(value)).filter((value) => Number.isFinite(value)).slice(0, 8)
  };
}

function lenovoRectShape(id, name, x, y, w, h, fillColor, lineColor) {
  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${ptToEmu(x)}" y="${ptToEmu(y)}"/><a:ext cx="${ptToEmu(w)}" cy="${ptToEmu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fillColor}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="${lineColor}"/></a:solidFill></a:ln></p:spPr>
</p:sp>`;
}

function ptToEmu(value) {
  return Math.round(Number(value || 0) * 12700);
}

function buildPptxFromTemplate(deck, templatePath) {
  const entries = readZipEntries(fs.readFileSync(templatePath));
  const contentSlideNumbers = [4, 5, 6, 7, 9, 10, 11, 13, 14, 15, 17, 18, 19];
  const sectionSlideNumbers = [3, 8, 12, 16];
  const sectionIndexes = [0, 4, 7, 10];

  return zipStore(
    entries.map(([name, data]) => {
      const match = name.match(/^ppt\/slides\/slide(\d+)\.xml$/);
      if (!match) {
        return [name, data];
      }

      const slideNumber = Number(match[1]);
      let xml = data.toString("utf8");

      if (slideNumber === 1) {
        const coverTitle = compactTitle(deck.title, 18);
        xml = replaceWholeShapeContaining(xml, "部门工作总结", coverTitle);
        xml = replaceFirstText(xml, "2O2X", extractYear(deck.title));
        xml = replaceFirstText(xml, "Q1", extractQuarter(deck.title));
        xml = resizeShapeContaining(xml, coverTitle, 3600);
      }

      const sectionPosition = sectionSlideNumbers.indexOf(slideNumber);
      if (sectionPosition >= 0) {
        const sectionSlide = deck.slides[sectionIndexes[sectionPosition]];
        if (sectionSlide?.title) {
          xml = replaceSectionHeading(xml, sectionSlide.title);
        }
      }

      const contentPosition = contentSlideNumbers.indexOf(slideNumber);
      if (contentPosition >= 0) {
        const slide = deck.slides[contentPosition];
        if (slide) {
          xml = fillTemplateContentSlide(xml, deck.title, slide, slideNumber);
          xml = resizeShapeContaining(xml, compactTitle(slide.title, 14), 2600);
        } else {
          xml = clearTemplatePlaceholderText(xml);
        }
      }

      if (slideNumber === 15) {
        xml = fillTemplateMetrics(xml, collectDeckMetrics(deck).slice(0, 3));
      }

      if (slideNumber === 17) {
        xml = fillTemplateMetrics(xml, collectDeckMetrics(deck).slice(3, 5));
      }

      xml = clearRemainingTemplateDemoText(xml);

      return [name, Buffer.from(xml, "utf8")];
    })
  );
}

function fillTemplateContentSlide(xml, deckTitle, slide, slideNumber) {
  let bulletIndex = 0;
  let headingReplaced = false;
  const pageLimits = { 4: 38, 5: 22, 6: 24, 7: 22, 9: 18, 10: 22, 11: 22, 13: 22, 14: 18, 15: 30, 17: 26, 18: 18, 19: 22 };
  const bulletLimit = pageLimits[slideNumber] || 32;

  return xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (full, encodedText) => {
    const text = decodeXml(encodedText).trim();

    if (/^FILL IN THE TEXT OF THE DOCUMENT TITLE HERE$/i.test(text)) {
      return `<a:t>${escapeXml(compactTitle(deckTitle, 22))}</a:t>`;
    }

    if (!headingReplaced && isSectionHeading(text)) {
      headingReplaced = true;
      return `<a:t>${escapeXml(compactTitle(slide.title, 14))}</a:t>`;
    }

    if (isTemplatePlaceholder(text)) {
      const replacement = compactBullet(slide.bullets[bulletIndex] || "", bulletLimit);
      bulletIndex += 1;
      return `<a:t>${escapeXml(replacement)}</a:t>`;
    }

    return full;
  });
}

function clearTemplatePlaceholderText(xml) {
  return xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (full, encodedText) => {
    const text = decodeXml(encodedText).trim();
    return isTemplatePlaceholder(text) ? "<a:t></a:t>" : full;
  });
}

function replaceSectionHeading(xml, title) {
  let replaced = false;
  return xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (full, encodedText) => {
    const text = decodeXml(encodedText).trim();
    if (!replaced && isSectionHeading(text)) {
      replaced = true;
      return `<a:t>${escapeXml(title)}</a:t>`;
    }
    return full;
  });
}

function replaceFirstText(xml, expected, replacement) {
  let replaced = false;
  return xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (full, encodedText) => {
    if (!replaced && decodeXml(encodedText).trim() === expected) {
      replaced = true;
      return `<a:t>${escapeXml(replacement)}</a:t>`;
    }
    return full;
  });
}

function isSectionHeading(text) {
  return /^(季度工作复盘总结|重点项目取得进展|暴露不足改进方案|下一阶段工作计划)$/.test(text);
}

function isTemplatePlaceholder(text) {
  return (
    /^(FILL IN THE TEXT|BUSINESS TITLE|WORK ANNUAL|REPORT|PROJECT NAME|TITLE\s*NAME|添加标题|请输入标题文字|请在此输入标题文字|标题文字内容)/i.test(text) ||
    (text.length > 5 &&
      /(请替换|请在此|在此输入|在此录入|点击输入|单击此处|添加具体内容|输入你的正文|文字内容|修改文字|复制您的文本|复制你的内容)/.test(text))
  );
}

function clearRemainingTemplateDemoText(xml) {
  const clearedShapes = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shape) => {
    const text = Array.from(shape.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g))
      .map((match) => decodeXml(match[1]).trim())
      .filter(Boolean)
      .join(" ");
    if (/^(YOUR CONTENT|PLEASE|ENTER THE RELEVANT|FILL IN THE TEXT)/i.test(text)) {
      return replaceAllShapeText(shape, "");
    }
    return shape;
  });

  return clearedShapes.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (full, encodedText) => {
    const text = decodeXml(encodedText).trim();
    const isDemo =
      /^(YOUR CONTENT|PLEASE|ENTER THE RELEVANT|FILL IN THE TEXT)/i.test(text) ||
      /^(BUSINESS|TITLE|ANNUAL|CORPORATE|WORK|SUMMARY|REPORT|PROJECT|NAME|2O2X TARGET|BUSINESS WORK SUMMARY|2O2X BUSINESS WORK SUMMARY|PPT模板|下载)$/i.test(text) ||
      /^(添加|标题|添加标题|请输入标题文字|请在此输入标题文字|标题文字内容|在此添加标题)$/.test(text);
    return isDemo ? "<a:t></a:t>" : full;
  });
}

function replaceWholeShapeContaining(xml, expected, replacement) {
  let replaced = false;
  return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shape) => {
    if (replaced || !shape.includes(escapeXml(expected))) return shape;
    replaced = true;
    return replaceAllShapeText(shape, replacement);
  });
}

function replaceAllShapeText(shape, replacement) {
  let first = true;
  return shape.replace(/<a:t>[\s\S]*?<\/a:t>/g, () => {
    const value = first ? replacement : "";
    first = false;
    return `<a:t>${escapeXml(value)}</a:t>`;
  });
}

function fillTemplateMetrics(xml, metrics) {
  let index = 0;
  return xml.replace(/<a:t>(37%|69%|88%|17\.8|22\.8)<\/a:t>/g, (full, original) => {
    const metric = metrics[index++];
    const suffix = original.includes("%") ? "%" : "";
    return `<a:t>${escapeXml(metric ? `${metric.value}${suffix}` : "")}</a:t>`;
  });
}

function collectDeckMetrics(deck) {
  if (Array.isArray(deck.metrics) && deck.metrics.length) return deck.metrics;
  const results = [];
  const seen = new Set();
  for (const slide of deck.slides || []) {
    for (const text of [slide.title, ...(slide.bullets || [])]) {
      const matches = String(text || "").matchAll(/(完成率|使用率|满意度|解决率|增长率|续约率|意向率|到场率|转化率|可用率)?[^，。；%]{0,10}?(\d+(?:\.\d+)?)%/g);
      for (const match of matches) {
        const value = Number(match[2]);
        if (value < 0 || value > 100 || seen.has(value)) continue;
        seen.add(value);
        results.push({ label: match[1] || "关键指标", value });
      }
    }
  }
  return results;
}

function compactTitle(value, maxLength) {
  const text = String(value || "").replace(/^#+\s*/, "").replace(/\s+/g, " ").trim();
  return text.length <= maxLength ? text : `${text.slice(0, maxLength - 1)}…`;
}

function compactBullet(value, maxLength) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  return text.length <= maxLength ? text : `${text.slice(0, maxLength - 1)}…`;
}

function extractYear(title) {
  return String(title || "").match(/20\d{2}/)?.[0] || String(new Date().getFullYear());
}

function extractQuarter(title) {
  const match = String(title || "").match(/第?([一二三四1234])季度/);
  const map = { 一: "Q1", 二: "Q2", 三: "Q3", 四: "Q4", 1: "Q1", 2: "Q2", 3: "Q3", 4: "Q4" };
  return map[match?.[1]] || "REPORT";
}

function resizeShapeContaining(xml, text, maxSize) {
  return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shape) => {
    if (!shape.includes(escapeXml(text))) return shape;
    return shape.replace(/\bsz="\d+"/g, `sz="${maxSize}"`);
  });
}

function decodeXml(value) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function readZipEntries(buffer) {
  const endSignature = 0x06054b50;
  let endOffset = -1;
  for (let offset = buffer.length - 22; offset >= Math.max(0, buffer.length - 65557); offset -= 1) {
    if (buffer.readUInt32LE(offset) === endSignature) {
      endOffset = offset;
      break;
    }
  }
  if (endOffset < 0) {
    throw new Error("PPT模板不是有效的PPTX文件。");
  }

  const entryCount = buffer.readUInt16LE(endOffset + 10);
  let centralOffset = buffer.readUInt32LE(endOffset + 16);
  const entries = [];

  for (let index = 0; index < entryCount; index += 1) {
    if (buffer.readUInt32LE(centralOffset) !== 0x02014b50) {
      throw new Error("PPT模板目录结构损坏。");
    }
    const method = buffer.readUInt16LE(centralOffset + 10);
    const compressedSize = buffer.readUInt32LE(centralOffset + 20);
    const fileNameLength = buffer.readUInt16LE(centralOffset + 28);
    const extraLength = buffer.readUInt16LE(centralOffset + 30);
    const commentLength = buffer.readUInt16LE(centralOffset + 32);
    const localOffset = buffer.readUInt32LE(centralOffset + 42);
    const name = buffer.slice(centralOffset + 46, centralOffset + 46 + fileNameLength).toString("utf8");
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = buffer.slice(dataOffset, dataOffset + compressedSize);
    let data;
    if (method === 0) data = compressed;
    else if (method === 8) data = zlib.inflateRawSync(compressed);
    else throw new Error(`PPT模板包含不支持的压缩方式：${method}`);
    entries.push([name, data]);
    centralOffset += 46 + fileNameLength + extraLength + commentLength;
  }

  return entries;
}

function buildPptx(deck) {
  const files = [];
  const slideCount = deck.slides.length;

  files.push(["[Content_Types].xml", contentTypesXml(slideCount)]);
  files.push(["_rels/.rels", rootRelsXml()]);
  files.push(["docProps/core.xml", coreXml(deck.title)]);
  files.push(["docProps/app.xml", appXml(slideCount)]);
  files.push(["ppt/presentation.xml", presentationXml(slideCount)]);
  files.push(["ppt/_rels/presentation.xml.rels", presentationRelsXml(slideCount)]);
  files.push(["ppt/theme/theme1.xml", themeXml()]);
  files.push(["ppt/slideMasters/slideMaster1.xml", slideMasterXml()]);
  files.push(["ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRelsXml()]);
  files.push(["ppt/slideLayouts/slideLayout1.xml", slideLayoutXml()]);

  deck.slides.forEach((slide, index) => {
    files.push([`ppt/slides/slide${index + 1}.xml`, slideXml(slide, index + 1)]);
    files.push([`ppt/slides/_rels/slide${index + 1}.xml.rels`, slideRelsXml()]);
  });

  return zipStore(files);
}

function contentTypesXml(slideCount) {
  const slideOverrides = Array.from({ length: slideCount }, (_, index) => {
    const slideNumber = index + 1;
    return `<Override PartName="/ppt/slides/slide${slideNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  ${slideOverrides}
</Types>`;
}

function rootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function coreXml(title) {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(title)}</dc:title>
  <dc:creator>Report Agent</dc:creator>
  <cp:lastModifiedBy>Report Agent</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

function appXml(slideCount) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Report Agent</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>${slideCount}</Slides>
</Properties>`;
}

function presentationXml(slideCount) {
  const slideIds = Array.from({ length: slideCount }, (_, index) => {
    const slideNumber = index + 1;
    return `<p:sldId id="${255 + slideNumber}" r:id="rId${slideNumber + 1}"/>`;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${slideIds}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
}

function presentationRelsXml(slideCount) {
  const slideRels = Array.from({ length: slideCount }, (_, index) => {
    const slideNumber = index + 1;
    return `<Relationship Id="rId${slideNumber + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNumber}.xml"/>`;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${slideRels}
</Relationships>`;
}

function themeXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Report Agent">
  <a:themeElements>
    <a:clrScheme name="ReportAgent">
      <a:dk1><a:srgbClr val="111827"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2><a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
      <a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="059669"/></a:accent2>
      <a:accent3><a:srgbClr val="DC2626"/></a:accent3><a:accent4><a:srgbClr val="7C3AED"/></a:accent4>
      <a:accent5><a:srgbClr val="EA580C"/></a:accent5><a:accent6><a:srgbClr val="0891B2"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink><a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="ReportAgent">
      <a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:majorFont>
      <a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="ReportAgent">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}

function slideMasterXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`;
}

function slideMasterRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;
}

function slideLayoutXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;
}

function slideRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
}

function slideXml(slide, slideNumber) {
  const subtitle = `由智能体生成 · 第 ${slideNumber} 页`;
  const bullets = slide.bullets.length > 0 ? slide.bullets : ["请补充更多内容。"];

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F8FAFC"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      ${textShape(2, "Title", slide.title, 685800, 520000, 10820400, 760000, 3200, "111827", true)}
      ${textShape(3, "Subtitle", subtitle, 685800, 1220000, 10820400, 360000, 1200, "64748B", false)}
      ${bulletShape(4, bullets, 900000, 1800000, 10300000, 3800000)}
      ${footerShape(5)}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function textShape(id, name, text, x, y, cx, cy, size, color, bold) {
  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" sz="${size}"${bold ? ' b="1"' : ""}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr><a:t>${escapeXml(text)}</a:t></a:r></a:p></p:txBody>
</p:sp>`;
}

function bulletShape(id, bullets, x, y, cx, cy) {
  const paragraphs = bullets
    .map(
      (bullet) => `<a:p>
  <a:pPr marL="342900" indent="-228600"><a:buChar char="•"/></a:pPr>
  <a:r><a:rPr lang="zh-CN" sz="1900"><a:solidFill><a:srgbClr val="1F2937"/></a:solidFill></a:rPr><a:t>${escapeXml(bullet)}</a:t></a:r>
</a:p>`
    )
    .join("");

  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="Bullets"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>${paragraphs}</p:txBody>
</p:sp>`;
}

function footerShape(id) {
  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="Footer"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="685800" y="6320000"/><a:ext cx="10820400" cy="70000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>
</p:sp>`;
}

function zipStore(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  entries.forEach(([name, content]) => {
    const nameBuffer = Buffer.from(name, "utf8");
    const data = Buffer.from(content, "utf8");
    const crc = crc32(data);

    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0x0800, 6);
    localHeader.writeUInt16LE(0, 8);
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(data.length, 18);
    localHeader.writeUInt32LE(data.length, 22);
    localHeader.writeUInt16LE(nameBuffer.length, 26);
    localHeader.writeUInt16LE(0, 28);

    localParts.push(localHeader, nameBuffer, data);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0x0800, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(data.length, 20);
    centralHeader.writeUInt32LE(data.length, 24);
    centralHeader.writeUInt16LE(nameBuffer.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);

    centralParts.push(centralHeader, nameBuffer);
    offset += localHeader.length + nameBuffer.length + data.length;
  });

  const centralDirectory = Buffer.concat(centralParts);
  const localFiles = Buffer.concat(localParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralDirectory.length, 12);
  end.writeUInt32LE(localFiles.length, 16);
  end.writeUInt16LE(0, 20);

  return Buffer.concat([localFiles, centralDirectory, end]);
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (let i = 0; i < buffer.length; i += 1) {
    crc ^= buffer[i];
    for (let j = 0; j < 8; j += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function sanitizeFileName(name) {
  return String(name || "智能体生成PPT")
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, "_")
    .slice(0, 60);
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function sendServerError(res, error) {
  console.error(error);
  res.status(error.statusCode || 500).json({
    success: false,
    message: error.message || "服务器内部错误，请查看终端日志。",
    detail: error.detail
  });
}

function startServer(port = PORT) {
  return app.listen(port, () => {
    console.log(`汇报智能体服务已启动：http://localhost:${port}`);
  });
}

if (require.main === module) {
  startServer();
}

module.exports = app;
module.exports.app = app;
module.exports.buildPptx = buildPptx;
module.exports.buildLenovoPptxFromTemplate = buildLenovoPptxFromTemplate;
module.exports.buildPptxFromTemplate = buildPptxFromTemplate;
module.exports.buildLocalPptDeck = buildLocalPptDeck;
module.exports.parseReportForTemplate = parseReportForTemplate;
module.exports.startServer = startServer;
