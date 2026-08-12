const pptxgen = require("pptxgenjs");

const pptx = new pptxgen();
pptx.defineLayout({ name: "XHS", width: 7.5, height: 10 });
pptx.layout = "XHS";
pptx.author = "nanfengy";
pptx.subject = "AI research visualization carousel";
pptx.title = "导师让我给组里讲AI作图";
pptx.company = "";
pptx.lang = "zh-CN";
pptx.theme = {
  headFontFace: "Microsoft YaHei",
  bodyFontFace: "Microsoft YaHei",
  lang: "zh-CN",
};
pptx.margin = 0;

const C = {
  navy: "173F70",
  blue: "2878C8",
  cyan: "22A7A0",
  green: "46A66B",
  coral: "E46F61",
  yellow: "F0B84B",
  ink: "17212B",
  gray: "66727D",
  pale: "EEF4F8",
  line: "D7E2EA",
  white: "FFFFFF",
  black: "111111",
};

const font = "Microsoft YaHei";

function addPageBase(slide, index, section) {
  slide.background = { color: C.white };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 7.5,
    h: 0.18,
    fill: { color: C.navy },
    line: { color: C.navy },
  });
  slide.addText(section, {
    x: 0.55,
    y: 0.42,
    w: 5.8,
    h: 0.28,
    fontFace: font,
    fontSize: 10,
    bold: true,
    color: C.blue,
    margin: 0,
  });
  slide.addText(String(index).padStart(2, "0"), {
    x: 6.38,
    y: 0.39,
    w: 0.55,
    h: 0.32,
    fontFace: "Arial",
    fontSize: 11,
    bold: true,
    align: "right",
    color: C.gray,
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 0.55,
    y: 9.43,
    w: 6.4,
    h: 0,
    line: { color: C.line, width: 1 },
  });
  slide.addText("nanfengy · AI 科研作图分享", {
    x: 0.55,
    y: 9.55,
    w: 5.4,
    h: 0.24,
    fontFace: font,
    fontSize: 8.5,
    color: C.gray,
    margin: 0,
  });
}

function addTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.55,
    y: 0.92,
    w: 6.4,
    h: 0.85,
    fontFace: font,
    fontSize: 27,
    bold: true,
    color: C.ink,
    breakLine: false,
    margin: 0,
    fit: "shrink",
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.55,
      y: 1.86,
      w: 6.25,
      h: 0.64,
      fontFace: font,
      fontSize: 13,
      color: C.gray,
      margin: 0,
      breakLine: false,
      fit: "shrink",
    });
  }
}

function addPill(slide, text, x, y, w, fill, color = C.white) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: 0.4,
    rectRadius: 0.06,
    fill: { color: fill },
    line: { color: fill },
  });
  slide.addText(text, {
    x,
    y: y + 0.04,
    w,
    h: 0.25,
    fontFace: font,
    fontSize: 9.5,
    bold: true,
    align: "center",
    color,
    margin: 0,
  });
}

function addCard(slide, x, y, w, h, title, body, accent = C.blue) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.06,
    fill: { color: C.white },
    line: { color: C.line, width: 1.1 },
    shadow: { type: "outer", color: "B8C6D1", blur: 1, angle: 45, distance: 1, opacity: 0.14 },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + 0.24,
    y: y + 0.27,
    w: 0.36,
    h: 0.36,
    fill: { color: accent },
    line: { color: accent },
  });
  slide.addText(title, {
    x: x + 0.72,
    y: y + 0.23,
    w: w - 0.94,
    h: 0.38,
    fontFace: font,
    fontSize: 14,
    bold: true,
    color: C.ink,
    margin: 0,
    fit: "shrink",
  });
  slide.addText(body, {
    x: x + 0.24,
    y: y + 0.82,
    w: w - 0.48,
    h: h - 1.05,
    fontFace: font,
    fontSize: 11,
    color: C.gray,
    valign: "mid",
    margin: 0,
    breakLine: false,
    fit: "shrink",
  });
}

function addConclusion(slide, text, color = C.navy) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: 8.42,
    w: 6.4,
    h: 0.68,
    rectRadius: 0.05,
    fill: { color },
    line: { color },
  });
  slide.addText(text, {
    x: 0.78,
    y: 8.61,
    w: 5.94,
    h: 0.28,
    fontFace: font,
    fontSize: 13,
    bold: true,
    align: "center",
    color: C.white,
    margin: 0,
    fit: "shrink",
  });
}

function addArrow(slide, x, y, w = 0.35) {
  slide.addShape(pptx.ShapeType.chevron, {
    x,
    y,
    w,
    h: 0.42,
    fill: { color: "BCD0DF" },
    line: { color: "BCD0DF" },
  });
}

// 1. Cover
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 7.5,
    h: 3.0,
    fill: { color: C.navy },
    line: { color: C.navy },
  });
  slide.addText("导师让我给组里讲", {
    x: 0.62,
    y: 0.82,
    w: 6.2,
    h: 0.62,
    fontFace: font,
    fontSize: 25,
    bold: true,
    color: C.white,
    margin: 0,
  });
  slide.addText("AI 作图", {
    x: 0.62,
    y: 1.48,
    w: 5.4,
    h: 0.92,
    fontFace: font,
    fontSize: 42,
    bold: true,
    color: C.white,
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.arc, {
    x: 5.25,
    y: 0.48,
    w: 1.55,
    h: 1.55,
    adjustPoint: 0.25,
    rotate: 12,
    line: { color: C.cyan, width: 6, transparency: 10 },
    fill: { color: C.navy, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.arc, {
    x: 5.62,
    y: 1.02,
    w: 1.25,
    h: 1.25,
    rotate: 205,
    line: { color: C.coral, width: 6, transparency: 5 },
    fill: { color: C.navy, transparency: 100 },
  });
  addPill(slide, "PPT", 0.62, 3.55, 1.05, C.blue);
  addPill(slide, "PDF", 1.82, 3.55, 1.05, C.cyan);
  addPill(slide, "工作流", 3.02, 3.55, 1.35, C.coral);
  slide.addText("博士生整理的 74 页科研作图分享", {
    x: 0.62,
    y: 4.28,
    w: 6.2,
    h: 0.72,
    fontFace: font,
    fontSize: 22,
    bold: true,
    color: C.ink,
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 0.62,
    y: 5.35,
    w: 5.9,
    h: 0,
    line: { color: C.line, width: 1.5 },
  });
  const steps = [
    ["01", "图形思维"],
    ["02", "技术路线"],
    ["03", "工具选型"],
    ["04", "完整工作流"],
  ];
  steps.forEach((item, i) => {
    const y = 5.78 + i * 0.72;
    slide.addText(item[0], {
      x: 0.68,
      y,
      w: 0.6,
      h: 0.35,
      fontFace: "Arial",
      fontSize: 12,
      bold: true,
      color: [C.blue, C.cyan, C.coral, C.yellow][i],
      margin: 0,
    });
    slide.addText(item[1], {
      x: 1.38,
      y: y - 0.01,
      w: 2.2,
      h: 0.36,
      fontFace: font,
      fontSize: 14,
      bold: true,
      color: C.ink,
      margin: 0,
    });
  });
  slide.addText("nanfengy", {
    x: 0.62,
    y: 9.25,
    w: 2,
    h: 0.3,
    fontFace: "Arial",
    fontSize: 10,
    bold: true,
    color: C.gray,
    margin: 0,
  });
}

// 2. Origin
{
  const slide = pptx.addSlide();
  addPageBase(slide, 2, "这件事是怎么开始的");
  addTitle(slide, "原本只想讲几个工具", "整理以后才发现，科研作图真正难的不是“生成”，而是“表达”。");
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: 2.85,
    w: 6.4,
    h: 1.55,
    rectRadius: 0.06,
    fill: { color: C.pale },
    line: { color: C.pale },
  });
  slide.addText("前段时间，导师邀请我给组里的师弟师妹做一次 AI 作图分享。", {
    x: 0.9,
    y: 3.17,
    w: 5.7,
    h: 0.85,
    fontFace: font,
    fontSize: 19,
    bold: true,
    color: C.navy,
    align: "center",
    valign: "mid",
    margin: 0,
    fit: "shrink",
  });
  addCard(slide, 0.55, 4.88, 3.02, 2.45, "最初的想法", "讲常用工具、提示词和几个案例，做成一次实用分享。", C.blue);
  addCard(slide, 3.93, 4.88, 3.02, 2.45, "最后的结果", "从图形思维讲到技术路线、工作流、质量判断和正式场景边界。", C.coral);
  addConclusion(slide, "工具只是后半程，图形思维才是起点。");
}

// 3. Four questions
{
  const slide = pptx.addSlide();
  addPageBase(slide, 3, "科研作图最难的 4 个问题");
  addTitle(slide, "真正难的，往往不是画", "开始做图之前，先把下面四个问题想明白。");
  const cards = [
    ["表达目标", "这张图到底想让读者理解或记住什么？", C.blue],
    ["图的类型", "它是流程图、架构图、机制图，还是概念图？", C.cyan],
    ["AI 的位置", "哪些环节适合交给 AI，哪些必须自己判断？", C.coral],
    ["交付形式", "如何变成可修改、可复用、能正式交付的文件？", C.yellow],
  ];
  cards.forEach((c, i) => {
    const x = i % 2 === 0 ? 0.55 : 3.93;
    const y = i < 2 ? 2.75 : 5.42;
    addCard(slide, x, y, 3.02, 2.25, c[0], c[1], c[2]);
  });
  addConclusion(slide, "先想清楚“做什么图”，再决定“怎么做”。", C.coral);
}

// 4. Six routes
{
  const slide = pptx.addSlide();
  addPageBase(slide, 4, "AI 作图不只有文生图");
  addTitle(slide, "六条路线，解决六类问题", "“输入一句提示词直接生图”，只是其中最显眼的一条路线。");
  const routes = [
    ["结构化", "逻辑稳定\n可编辑", C.blue],
    ["代码化", "可复用\n可维护", C.cyan],
    ["生成式", "视觉探索\n快速发散", C.coral],
    ["可控工作流", "细节控制\n结果一致", C.green],
    ["可编辑化", "继续修改\n正式交付", C.yellow],
    ["视频化", "动态叙事\n时间表达", "7656A8"],
  ];
  routes.forEach((r, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.55 : 3.93;
    const y = 2.55 + row * 1.82;
    addCard(slide, x, y, 3.02, 1.52, r[0], r[1], r[2]);
  });
  addConclusion(slide, "没有一种工具能包办全部，关键是路线协同。");
}

// 5. Types
{
  const slide = pptx.addSlide();
  addPageBase(slide, 5, "先判断图类型");
  addTitle(slide, "表达任务不同，图的类型就不同", "图类型不是由外观决定，而是由你要表达的内容决定。");
  const items = [
    ["讲结构", "架构图 / 系统图", C.blue],
    ["讲步骤", "流程图", C.cyan],
    ["讲原理", "机制示意图", C.coral],
    ["讲方法", "框架图 / 技术路线图", C.green],
    ["讲差异", "对比图", C.yellow],
    ["讲氛围", "封面图 / 概念图", "7656A8"],
    ["讲过程", "动态展示", "357A78"],
  ];
  items.forEach((item, i) => {
    const y = 2.52 + i * 0.72;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.55,
      y,
      w: 6.4,
      h: 0.55,
      rectRadius: 0.04,
      fill: { color: i % 2 === 0 ? "F7FAFC" : C.white },
      line: { color: C.line, width: 0.8 },
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 0.8,
      y: y + 0.12,
      w: 0.28,
      h: 0.28,
      fill: { color: item[2] },
      line: { color: item[2] },
    });
    slide.addText(item[0], {
      x: 1.28,
      y: y + 0.12,
      w: 1.25,
      h: 0.28,
      fontFace: font,
      fontSize: 12.5,
      bold: true,
      color: C.ink,
      margin: 0,
    });
    slide.addText(item[1], {
      x: 2.7,
      y: y + 0.12,
      w: 3.75,
      h: 0.28,
      fontFace: font,
      fontSize: 12,
      color: C.gray,
      margin: 0,
      fit: "shrink",
    });
  });
  addConclusion(slide, "目标决定类型，类型决定路线。", C.cyan);
}

// 6. Workflow
{
  const slide = pptx.addSlide();
  addPageBase(slide, 6, "完整工作流");
  addTitle(slide, "高质量科研图，是逐步构建出来的", "生成只是其中的一步，前期定义和后期校验同样重要。");
  const stages = [
    ["1", "明确边界", C.blue],
    ["2", "逻辑分解", C.cyan],
    ["3", "判断类型", C.green],
    ["4", "路线与工具", C.yellow],
    ["5", "生成初稿", C.coral],
    ["6", "精修交付", "7656A8"],
    ["7", "校验复用", C.navy],
  ];
  stages.forEach((s, i) => {
    const y = 2.52 + i * 0.73;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 0.68,
      y,
      w: 0.48,
      h: 0.48,
      fill: { color: s[2] },
      line: { color: s[2] },
    });
    slide.addText(s[0], {
      x: 0.68,
      y: y + 0.09,
      w: 0.48,
      h: 0.22,
      fontFace: "Arial",
      fontSize: 10.5,
      bold: true,
      align: "center",
      color: C.white,
      margin: 0,
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 1.42,
      y: y - 0.02,
      w: 4.95,
      h: 0.52,
      rectRadius: 0.04,
      fill: { color: "F7FAFC" },
      line: { color: C.line, width: 0.9 },
    });
    slide.addText(s[1], {
      x: 1.7,
      y: y + 0.1,
      w: 4.35,
      h: 0.25,
      fontFace: font,
      fontSize: 13,
      bold: true,
      color: C.ink,
      margin: 0,
    });
    if (i < stages.length - 1) {
      slide.addShape(pptx.ShapeType.line, {
        x: 0.92,
        y: y + 0.48,
        w: 0,
        h: 0.28,
        line: { color: "B9C9D6", width: 2, beginArrowType: "none", endArrowType: "triangle" },
      });
    }
  });
  addConclusion(slide, "先把逻辑立住，再处理视觉细节。");
}

// 7. Tool chain
{
  const slide = pptx.addSlide();
  addPageBase(slide, 7, "一套常用的工具组合");
  addTitle(slide, "不是找一个“最强工具”", "成熟的做法通常是让不同工具负责不同环节。");
  const blocks = [
    ["大模型", "梳理模块、层级和关系", C.blue],
    ["Draw.io\nMermaid", "搭建逻辑骨架", C.cyan],
    ["生成式工具", "探索构图和视觉方向", C.coral],
    ["PPT / SVG", "精修并可编辑交付", C.green],
  ];
  blocks.forEach((b, i) => {
    const y = 2.65 + i * 1.32;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.7,
      y,
      w: 2.02,
      h: 0.9,
      rectRadius: 0.05,
      fill: { color: b[2] },
      line: { color: b[2] },
    });
    slide.addText(b[0], {
      x: 0.88,
      y: y + 0.19,
      w: 1.66,
      h: 0.48,
      fontFace: font,
      fontSize: 13,
      bold: true,
      align: "center",
      valign: "mid",
      color: C.white,
      margin: 0,
      fit: "shrink",
    });
    addArrow(slide, 2.92, y + 0.23, 0.4);
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 3.55,
      y,
      w: 3.18,
      h: 0.9,
      rectRadius: 0.05,
      fill: { color: "F7FAFC" },
      line: { color: C.line, width: 1 },
    });
    slide.addText(b[1], {
      x: 3.82,
      y: y + 0.22,
      w: 2.64,
      h: 0.4,
      fontFace: font,
      fontSize: 13,
      bold: true,
      color: C.ink,
      margin: 0,
      fit: "shrink",
    });
  });
  addConclusion(slide, "高效不在于工具数量，在于路线组合与协同。", C.coral);
}

// 8. Checklist
{
  const slide = pptx.addSlide();
  addPageBase(slide, 8, "正式科研场景一定要检查");
  addTitle(slide, "“看起来不错”不等于“表达有效”", "越正式的场景，越不能把 AI 结果直接拿来使用。");
  const checks = [
    ["结构关系", "模块、层级和箭头是否正确", C.blue],
    ["文字数字", "标签、公式和数据是否准确", C.cyan],
    ["专业细节", "对象与机制是否失真", C.coral],
    ["视觉一致", "颜色、字体、线条是否统一", C.green],
    ["交付可靠", "文件是否可编辑、可复用", C.yellow],
    ["合规边界", "版权、隐私和未公开数据", "7656A8"],
  ];
  checks.forEach((c, i) => {
    const x = i % 2 === 0 ? 0.55 : 3.93;
    const y = 2.56 + Math.floor(i / 2) * 1.84;
    addCard(slide, x, y, 3.02, 1.5, c[0], c[1], c[2]);
  });
  addConclusion(slide, "AI 负责提效，研究者负责最终判断。", C.navy);
}

// 9. Closing
{
  const slide = pptx.addSlide();
  addPageBase(slide, 9, "资料说明");
  addTitle(slide, "这次整理，我准备放出这些内容", "公开版已经做了脱敏和轻量化处理。");
  const list = [
    ["74 页", "公开版 PDF", C.blue],
    ["可编辑", "轻量版 PPT", C.cyan],
    ["可复用", "科研作图工作流", C.green],
    ["可执行", "质量检查框架", C.coral],
    ["有边界", "案例与使用规范", C.yellow],
  ];
  list.forEach((item, i) => {
    const y = 2.55 + i * 0.9;
    addPill(slide, item[0], 0.62, y, 1.25, item[2]);
    slide.addText(item[1], {
      x: 2.15,
      y: y + 0.06,
      w: 4.35,
      h: 0.32,
      fontFace: font,
      fontSize: 15,
      bold: true,
      color: C.ink,
      margin: 0,
    });
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.62,
    y: 7.42,
    w: 6.25,
    h: 1.2,
    rectRadius: 0.06,
    fill: { color: C.pale },
    line: { color: C.pale },
  });
  slide.addText("你最想解决哪一种科研作图问题？", {
    x: 0.92,
    y: 7.73,
    w: 5.65,
    h: 0.42,
    fontFace: font,
    fontSize: 18,
    bold: true,
    align: "center",
    color: C.navy,
    margin: 0,
  });
  slide.addText("GitHub：发布时补充", {
    x: 0.62,
    y: 9.0,
    w: 3.2,
    h: 0.26,
    fontFace: font,
    fontSize: 9.5,
    color: C.gray,
    margin: 0,
  });
}

pptx.writeFile({ fileName: "xiaohongshu/AI科研作图-小红书9页.pptx" });
