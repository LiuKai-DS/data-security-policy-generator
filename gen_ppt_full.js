const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';
pptx.title = '以人为本的数据安全运营方法 - WorkBuddy Skill';
pptx.author = 'Law42Pulse';

// ========== 配色方案 ==========
const COLORS = {
  darkBg: '0F172A',       // 深色背景
  primary: '3B82F6',      // 主蓝
  accent: 'F43F5E',       // 玫瑰红强调
  success: '10B981',      // 绿色
  warning: 'F59E0B',      // 橙色
  purple: '8B5CF6',       // 紫色
  teal: '14B8A6',         // 青色
  orange: 'F97316',       // 橙色
  white: 'FFFFFF',
  lightGray: 'F1F5F9',
  darkGray: '334155',
  midGray: '64748B',
  border: 'CBD5E1',
};

// ========== 辅助函数 ==========
function addSlideNum(slide, num, total) {
  slide.addText(`${num} / ${total}`, {
    x: 9.2, y: 5.2, w: 0.6, h: 0.3,
    fontSize: 8, color: COLORS.midGray, align: 'right'
  });
}

// 阴影工厂
const makeShadow = () => ({
  type: 'outer', color: '000000',
  blur: 8, offset: 3, angle: 135, opacity: 0.12
});

const TOTAL = 22;
let slideNum = 0;

// ========== SLIDE 1: 封面 ==========
slideNum++;
let s1 = pptx.addSlide();
s1.background = { color: COLORS.darkBg };

// 装饰圆形
s1.addShape(pptx.shapes.OVAL, {
  x: 7.5, y: -1, w: 4, h: 4,
  fill: { color: COLORS.primary, transparency: 85 }
});
s1.addShape(pptx.shapes.OVAL, {
  x: -1, y: 3.5, w: 3, h: 3,
  fill: { color: COLORS.accent, transparency: 85 }
});

// 产品标识
s1.addText('WorkBuddy Skill', {
  x: 0.5, y: 1.5, w: 9, h: 0.5,
  fontSize: 14, color: COLORS.primary, bold: true,
  charSpacing: 8
});

// 主标题
s1.addText('以人为本的\n数据安全运营方法', {
  x: 0.5, y: 2.0, w: 9, h: 1.8,
  fontSize: 44, color: COLORS.white, bold: true,
  lineSpacing: 52
});

// 副标题
s1.addText('基于《央国企总部数据安全策略设计》的工程化实现', {
  x: 0.5, y: 4.0, w: 9, h: 0.5,
  fontSize: 16, color: COLORS.midGray
});

// 版本标签
s1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.7, w: 1.5, h: 0.4,
  fill: { color: COLORS.accent }, rectRadius: 0.05
});
s1.addText('v2.1', {
  x: 0.5, y: 4.7, w: 1.5, h: 0.4,
  fontSize: 14, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
});

// GitHub
s1.addText('github.com/LiuKai-DS/data-security-policy-generator', {
  x: 0.5, y: 5.2, w: 9, h: 0.3,
  fontSize: 11, color: COLORS.midGray
});

addSlideNum(s1, slideNum, TOTAL);

// ========== SLIDE 2: 核心理念 ==========
slideNum++;
let s2 = pptx.addSlide();
s2.background = { color: COLORS.white };

s2.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.primary }
});

s2.addText('核心理念', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s2.addText('策略是死的，人是活的', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

// 核心观点卡片
const corePoints = [
  { icon: '01', title: '策略 ≠ 安全', desc: '策略只是工具，运营才是核心。再好的策略不运营就是废纸。', color: COLORS.accent },
  { icon: '02', title: '以人为本', desc: '从人的行为出发设计策略，不是从数据出发。先管住人，再管数据。', color: COLORS.primary },
  { icon: '03', title: '三阶段递进', desc: '前期建基线 → 中期精细化 → 后期精准化。阶段不同，策略不同。', color: COLORS.success },
  { icon: '04', title: '风险偏好驱动', desc: '不同管控风格（激进/平衡/保守）决定策略动作上限，策略必须适配组织文化。', color: COLORS.purple },
];

corePoints.forEach((p, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.7;
  const y = 1.5 + row * 1.9;

  // 卡片背景
  s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 4.3, h: 1.6,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.1
  });

  // 序号圆
  s2.addShape(pptx.shapes.OVAL, {
    x: x + 0.15, y: y + 0.15, w: 0.5, h: 0.5,
    fill: { color: p.color }
  });
  s2.addText(p.icon, {
    x: x + 0.15, y: y + 0.15, w: 0.5, h: 0.5,
    fontSize: 12, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });

  // 标题
  s2.addText(p.title, {
    x: x + 0.8, y: y + 0.2, w: 3.3, h: 0.4,
    fontSize: 16, color: COLORS.darkBg, bold: true
  });

  // 描述
  s2.addText(p.desc, {
    x: x + 0.2, y: y + 0.7, w: 3.9, h: 0.8,
    fontSize: 12, color: COLORS.darkGray
  });
});

addSlideNum(s2, slideNum, TOTAL);

// ========== SLIDE 3: 三阶段框架总览 ==========
slideNum++;
let s3 = pptx.addSlide();
s3.background = { color: COLORS.white };

s3.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.primary }
});

s3.addText('三阶段策略设计框架', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s3.addText('粗 → 细 → 精准，逐步收敛数据安全风险', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

// 三个阶段
const stages = [
  {
    num: '第一阶段',
    title: '行为管控策略',
    subtitle: '快速摸底',
    icon: '🎯',
    desc: '只要行为是恶意的，\n就视为有风险',
    strategies: '331 条通用策略',
    types: 'R/L/A/C/X/S 六类',
   适用: '企业无防护 / 项目前期 / 快速见效',
    color: COLORS.accent,
    arrow: '→'
  },
  {
    num: '第二阶段',
    title: '业务数据策略',
    subtitle: '精细化运营',
    icon: '🔍',
    desc: '提炼企业实际业务数据，\n针对性监控',
    strategies: '13 个部门关键字',
    types: '制药/半导体/能源/金融等',
   适用: '已有基础管控 / 需精细化',
    color: COLORS.primary,
    arrow: '→'
  },
  {
    num: '第三阶段',
    title: '融合策略',
    subtitle: '精准定位',
    icon: '⚡',
    desc: '行为+内容都为敏感，\n才视为违规',
    strategies: '200 条融合规则',
    types: '行为×内容双维度',
   适用: '需找真实泄露源 / 运营机制完善',
    color: COLORS.success,
    arrow: ''
  },
];

stages.forEach((s, i) => {
  const x = 0.4 + i * 3.2;
  const y = 1.4;
  const w = 3.0;

  // 卡片
  s3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h: 3.9,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.1
  });

  // 顶部色条
  s3.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h: 0.15,
    fill: { color: s.color }
  });

  // 阶段标签
  s3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.1, y: y + 0.25, w: 1.4, h: 0.35,
    fill: { color: s.color }, rectRadius: 0.05
  });
  s3.addText(s.num, {
    x: x + 0.1, y: y + 0.25, w: 1.4, h: 0.35,
    fontSize: 10, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });

  // 标题
  s3.addText(s.title, {
    x: x + 0.1, y: y + 0.7, w: 2.8, h: 0.45,
    fontSize: 16, color: COLORS.darkBg, bold: true
  });

  // 副标题
  s3.addText(s.subtitle, {
    x: x + 0.1, y: y + 1.1, w: 2.8, h: 0.3,
    fontSize: 12, color: s.color, bold: true
  });

  // 描述
  s3.addText(s.desc, {
    x: x + 0.1, y: y + 1.45, w: 2.8, h: 0.8,
    fontSize: 11, color: COLORS.darkGray
  });

  // 数据标签
  const dataLines = [
    s.strategies, s.types, s.适用
  ];
  dataLines.forEach((line, li) => {
    s3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.1, y: y + 2.4 + li * 0.5, w: 2.8, h: 0.4,
      fill: { color: COLORS.white }, rectRadius: 0.05
    });
    s3.addText(line, {
      x: x + 0.1, y: y + 2.4 + li * 0.5, w: 2.8, h: 0.4,
      fontSize: 9, color: COLORS.darkGray, align: 'center', valign: 'middle'
    });
  });

  // 箭头
  if (s.arrow) {
    s3.addText(s.arrow, {
      x: x + w, y: y + 1.8, w: 0.2, h: 0.5,
      fontSize: 20, color: COLORS.midGray, align: 'center', valign: 'middle'
    });
  }
});

addSlideNum(s3, slideNum, TOTAL);

// ========== SLIDE 4: 策略类型详解 ==========
slideNum++;
let s4 = pptx.addSlide();
s4.background = { color: COLORS.white };

s4.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.primary }
});

s4.addText('第一阶段：6 种策略类型详解', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s4.addText('331 条通用行为策略，覆盖所有常见数据泄露路径', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

const strategyTypes = [
  { type: 'R', name: '规则型', desc: '时间/频率/通道条件\n基于行为特征的规则触发', examples: '非工作时间外发\n异常频率外发\n私发个人邮箱', color: COLORS.accent },
  { type: 'L', name: '限制型', desc: '软件/设备/协议检测\n禁止特定工具或通道', examples: 'VPN软件检测\n代理工具检测\n远程控制软件', color: COLORS.primary },
  { type: 'A', name: '访问型', desc: '权限/批量行为检测\n异常访问模式识别', examples: '批量下载检测\n权限外数据访问\n共享账号检测', color: COLORS.success },
  { type: 'C', name: '内容型', desc: '关键字/正则/文件类型\n数据内容直接匹配', examples: '身份证号正则\n银行卡号正则\n邮箱地址正则', color: COLORS.purple },
  { type: 'X', name: '上下文型', desc: '组合条件\n多字段联合判断', examples: '身份证+姓名\n手机号+地址\n银行卡+持卡人', color: COLORS.warning },
  { type: 'S', name: '敏感行为型', desc: '规避检测\n试图绕过管控的行为', examples: '压缩深度检测\n扩展名伪装\n隐藏文件发送', color: COLORS.teal },
];

strategyTypes.forEach((t, i) => {
  const col = i % 3;
  const row = Math.floor(i / 3);
  const x = 0.4 + col * 3.15;
  const y = 1.4 + row * 2.1;

  // 卡片
  s4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 2.95, h: 1.9,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.08
  });

  // 左侧色条
  s4.addShape(pptx.shapes.RECTANGLE, {
    x, y, w: 0.1, h: 1.9,
    fill: { color: t.color }
  });

  // 类型标识
  s4.addShape(pptx.shapes.OVAL, {
    x: x + 0.2, y: y + 0.12, w: 0.55, h: 0.55,
    fill: { color: t.color }
  });
  s4.addText(t.type, {
    x: x + 0.2, y: y + 0.12, w: 0.55, h: 0.55,
    fontSize: 18, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });

  // 名称
  s4.addText(t.name, {
    x: x + 0.85, y: y + 0.18, w: 1.9, h: 0.35,
    fontSize: 14, color: COLORS.darkBg, bold: true
  });

  // 描述
  s4.addText(t.desc, {
    x: x + 0.85, y: y + 0.5, w: 2.0, h: 0.55,
    fontSize: 9, color: COLORS.midGray
  });

  // 示例
  s4.addText('示例：', {
    x: x + 0.15, y: y + 1.1, w: 0.5, h: 0.25,
    fontSize: 9, color: COLORS.darkGray, bold: true
  });
  s4.addText(t.examples, {
    x: x + 0.15, y: y + 1.35, w: 2.65, h: 0.5,
    fontSize: 9, color: COLORS.darkGray
  });
});

addSlideNum(s4, slideNum, TOTAL);

// ========== SLIDE 5: 策略示例：规则型 ==========
slideNum++;
let s5 = pptx.addSlide();
s5.background = { color: COLORS.white };

s5.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.accent }
});

s5.addText('策略示例：规则型（R）行为管控', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s5.addText('基于时间、频率、通道等条件触发，不区分数据类型', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

// 示例表格
const rExamples = [
  { id: 'R001', name: '非工作时间外发敏感文件', cond: '22:00 - 06:00', dept: '通用', data: 'L4/L3', action: '阻断（激进）/ 审批（平衡）', legal: '《劳动法》工时划分' },
  { id: 'R002', name: '异常频率外发', cond: '30min内外发>10次', dept: '通用', data: 'L4', action: '阻断（激进）/ 审批（平衡）', legal: '《数据安全法》第21条' },
  { id: 'R003', name: '非常用通道传输', cond: '非白名单IM工具', dept: '通用', data: 'L4/L3', action: '阻断 / 审批+告警', legal: '《网络安全法》第27条' },
  { id: 'R004', name: '私发个人邮箱', cond: '发送至个人邮箱', dept: '通用', data: 'L4', action: '阻断（激进）/ 审批（平衡）', legal: '《个人信息保护法》第29条' },
  { id: 'R005', name: '外部云盘上传', cond: '未授权云盘URL', dept: '通用', data: 'L4/L3', action: '阻断 / 审批', legal: '《数据安全法》第31条' },
  { id: 'R011', name: '频繁邮件外发', cond: '>20次/小时', dept: '采购/销售', data: 'L3', action: '审批+告警', legal: '制造业业务基线' },
];

// 表头
const tableHeaders = ['编号', '场景名称', '触发条件', '适用部门', '数据级别', '建议动作', '法规依据'];
const colWidths = [0.7, 1.8, 1.3, 0.9, 0.8, 1.5, 1.2];
let tx = 0.3;
tableHeaders.forEach((h, i) => {
  s5.addShape(pptx.shapes.RECTANGLE, {
    x: tx, y: 1.3, w: colWidths[i], h: 0.35,
    fill: { color: COLORS.accent }
  });
  s5.addText(h, {
    x: tx, y: 1.3, w: colWidths[i], h: 0.35,
    fontSize: 9, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });
  tx += colWidths[i];
});

// 数据行
rExamples.forEach((row, ri) => {
  const ry = 1.65 + ri * 0.6;
  const bgColor = ri % 2 === 0 ? COLORS.lightGray : COLORS.white;
  const vals = [row.id, row.name, row.cond, row.dept, row.data, row.action, row.legal];
  let cx = 0.3;
  vals.forEach((v, ci) => {
    s5.addShape(pptx.shapes.RECTANGLE, {
      x: cx, y: ry, w: colWidths[ci], h: 0.55,
      fill: { color: bgColor },
      line: { color: COLORS.border, width: 0.5 }
    });
    const fontColor = ci === 0 ? COLORS.accent : (ci === 1 ? COLORS.darkBg : COLORS.darkGray);
    const bold = ci === 0;
    s5.addText(v, {
      x: cx + 0.05, y: ry, w: colWidths[ci] - 0.1, h: 0.55,
      fontSize: 8, color: fontColor, bold, align: 'center', valign: 'middle'
    });
    cx += colWidths[ci];
  });
});

addSlideNum(s5, slideNum, TOTAL);

// ========== SLIDE 6: 内容型策略 ==========
slideNum++;
let s6 = pptx.addSlide();
s6.background = { color: COLORS.white };

s6.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.purple }
});

s6.addText('策略示例：内容型（C）敏感信息检测', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s6.addText('基于正则表达式匹配，直接识别敏感数据类型', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

const cExamples = [
  { id: 'C001', name: '身份证号', regex: '[1-9]\\d{5}(?:19|20)\\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\\d|3[01])\\d{3}[\\dXx]', level: 'L3', legal: '《个人信息保护法》第28条' },
  { id: 'C002', name: '手机号码', regex: '1[3-9]\\d{9}', level: 'L3', legal: '《个人信息保护法》第28条' },
  { id: 'C003', name: '银行卡号', regex: '[1-9]\\d{12,18}', level: 'L3', legal: '《个人信息保护法》第28条' },
  { id: 'C004', name: '邮箱地址', regex: '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}', level: 'L2', legal: '《个人信息保护法》' },
  { id: 'C005', name: '护照号码', regex: '[EeGg]\\d{8}', level: 'L3', legal: '《个人信息保护法》第28条' },
  { id: 'C006', name: '军官证号', regex: '[A-Z]\\d{7,9}', level: 'L4', legal: '《个人信息保护法》第28条' },
  { id: 'C007', name: '台胞证号', regex: '\\d{8,10}', level: 'L3', legal: '《个人信息保护法》第28条' },
  { id: 'C008', name: '营业执照', regex: '9\\d{13}|\\d{15}', level: 'L2', legal: '《数据安全法》' },
];

// 表格
const cHeaders = ['编号', '数据类型', '正则表达式', '数据级别', '法规依据'];
const cColW = [0.6, 1.1, 5.2, 0.9, 1.4];
let cTx = 0.3;
cHeaders.forEach((h, i) => {
  s6.addShape(pptx.shapes.RECTANGLE, {
    x: cTx, y: 1.3, w: cColW[i], h: 0.35,
    fill: { color: COLORS.purple }
  });
  s6.addText(h, {
    x: cTx, y: 1.3, w: cColW[i], h: 0.35,
    fontSize: 9, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });
  cTx += cColW[i];
});

cExamples.forEach((row, ri) => {
  const ry = 1.65 + ri * 0.48;
  const bgColor = ri % 2 === 0 ? COLORS.lightGray : COLORS.white;
  const vals = [row.id, row.name, row.regex, row.level, row.legal];
  let cx = 0.3;
  vals.forEach((v, ci) => {
    s6.addShape(pptx.shapes.RECTANGLE, {
      x: cx, y: ry, w: cColW[ci], h: 0.45,
      fill: { color: bgColor },
      line: { color: COLORS.border, width: 0.5 }
    });
    const fontColor = ci === 3 ? COLORS.warning : COLORS.darkGray;
    const bold = ci === 0;
    s6.addText(v, {
      x: cx + 0.05, y: ry, w: cColW[ci] - 0.1, h: 0.45,
      fontSize: ci === 2 ? 7 : 8, color: fontColor, bold, align: 'center', valign: 'middle'
    });
    cx += cColW[ci];
  });
});

addSlideNum(s6, slideNum, TOTAL);

// ========== SLIDE 7: 部门关键字 ==========
slideNum++;
let s7 = pptx.addSlide();
s7.background = { color: COLORS.white };

s7.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.success }
});

s7.addText('第二阶段：13 个部门业务数据关键字', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s7.addText('根据行业类型和企业实际业务，提炼针对性监控关键字', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

const departments = [
  { name: '研发部', l4: '源代码/芯片设计文件/工艺配方', l3: '专利申请文件/测试数据', color: COLORS.accent },
  { name: '采购部', l4: '供应商底价/评标底价', l3: '供应商名单/采购方案', color: COLORS.primary },
  { name: '财务部', l4: '高管薪酬/财务预算', l3: '财务报告/审计报告', color: COLORS.success },
  { name: '销售部', l4: '客户资产组合/投资策略', l3: '客户名单/销售渠道', color: COLORS.purple },
  { name: '生产部', l4: '核心工艺/配方参数', l3: '生产计划/物料清单', color: COLORS.warning },
  { name: '人力资源', l4: '高管信息/薪酬结构', l3: '员工信息/考核数据', color: COLORS.teal },
  { name: '战略发展', l4: '战略规划/投决记录', l3: '行业分析报告', color: COLORS.orange },
  { name: '信息化部', l4: '系统架构/核心源码', l3: '运维手册/拓扑图', color: COLORS.darkGray },
  { name: '法务部', l4: '核心合同条款/仲裁方案', l3: '合同模板/法规解读', color: COLORS.accent },
  { name: '市场部', l4: '定价策略/竞争分析', l3: '市场调研报告', color: COLORS.primary },
  { name: '质量管理', l4: '核心质量标准', l3: '检测方法/质量报告', color: COLORS.success },
  { name: '供应链', l4: '核心供应商关系', l3: '物流数据/库存策略', color: COLORS.purple },
  { name: '行政办公室', l4: '党委决议/人事安排', l3: '公章使用记录', color: COLORS.warning },
];

departments.forEach((d, i) => {
  const col = i % 3;
  const row = Math.floor(i / 3);
  const x = 0.3 + col * 3.2;
  const y = 1.35 + row * 1.05;

  // 卡片
  s7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 3.1, h: 0.9,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.06
  });

  // 左侧色条
  s7.addShape(pptx.shapes.RECTANGLE, {
    x, y, w: 0.1, h: 0.9,
    fill: { color: d.color }
  });

  // 部门名称
  s7.addText(d.name, {
    x: x + 0.18, y: y + 0.06, w: 2.8, h: 0.3,
    fontSize: 11, color: COLORS.darkBg, bold: true
  });

  // L4
  s7.addText(`L4：${d.l4}`, {
    x: x + 0.18, y: y + 0.38, w: 2.8, h: 0.25,
    fontSize: 8, color: COLORS.accent
  });

  // L3
  s7.addText(`L3：${d.l3}`, {
    x: x + 0.18, y: y + 0.62, w: 2.8, h: 0.25,
    fontSize: 8, color: COLORS.darkGray
  });
});

addSlideNum(s7, slideNum, TOTAL);

// ========== SLIDE 8: 融合策略 ==========
slideNum++;
let s8 = pptx.addSlide();
s8.background = { color: COLORS.white };

s8.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.warning }
});

s8.addText('第三阶段：行为 × 内容融合策略', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s8.addText('双维度精准定位，行为 + 内容都为敏感才触发，误报率最低', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

// 融合矩阵示意
s8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 1.4, w: 9, h: 2.8,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.1
});

s8.addText('融合策略矩阵示意', {
  x: 0.7, y: 1.5, w: 3, h: 0.3,
  fontSize: 12, color: COLORS.darkBg, bold: true
});

// 矩阵
const matW = 2.5, matH = 1.6;
const matX = 1.0, matY = 1.9;
const matColor = [
  [COLORS.warning, COLORS.accent, COLORS.accent],
  [COLORS.success, COLORS.warning, COLORS.accent],
  [COLORS.success, COLORS.success, COLORS.warning],
];

// 行标签
const rowLabels = ['L4 核心商密', 'L3 普通商密', 'L2 内部知悉'];
rowLabels.forEach((l, ri) => {
  s8.addText(l, {
    x: 0.5, y: matY + ri * (matH / 3) + 0.1, w: 0.5, h: 0.5,
    fontSize: 8, color: COLORS.darkGray, rotate: 270
  });
});

// 列标签
const colLabels = ['非工作时间行为', '频繁外发行为', '私发邮箱行为'];
colLabels.forEach((l, ci) => {
  s8.addText(l, {
    x: matX + ci * matW + 0.3, y: matY - 0.3, w: matW - 0.1, h: 0.25,
    fontSize: 8, color: COLORS.darkGray, align: 'center'
  });
});

for (let ri = 0; ri < 3; ri++) {
  for (let ci = 0; ci < 3; ci++) {
    const cx = matX + ci * matW;
    const cy = matY + ri * (matH / 3);

    s8.addShape(pptx.shapes.RECTANGLE, {
      x: cx, y: cy, w: matW - 0.05, h: matH / 3 - 0.05,
      fill: { color: matColor[ri][ci], transparency: 70 }
    });

    const label = ri === 0 && ci >= 0 ? '阻断' : (ri === 2 && ci === 2 ? '弹窗' : '审批');
    s8.addText(label, {
      x: cx, y: cy, w: matW - 0.05, h: matH / 3 - 0.05,
      fontSize: 10, color: COLORS.darkBg, bold: true, align: 'center', valign: 'middle'
    });
  }
}

// 图例
const legends = [
  { label: '阻断', color: COLORS.accent },
  { label: '审批', color: COLORS.warning },
  { label: '弹窗', color: COLORS.success },
];
legends.forEach((l, i) => {
  s8.addShape(pptx.shapes.RECTANGLE, {
    x: 7.5 + i * 1.1, y: 2.2, w: 0.3, h: 0.3,
    fill: { color: l.color, transparency: 70 }
  });
  s8.addText(l.label, {
    x: 7.85 + i * 1.1, y: 2.2, w: 0.7, h: 0.3,
    fontSize: 9, color: COLORS.darkGray, valign: 'middle'
  });
});

// 右侧说明
s8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 6.8, y: 2.7, w: 2.7, h: 1.3,
  fill: { color: COLORS.white },
  rectRadius: 0.06
});

s8.addText('融合策略优势', {
  x: 6.9, y: 2.8, w: 2.5, h: 0.3,
  fontSize: 11, color: COLORS.darkBg, bold: true
});

s8.addText('• 精准度最高\n• 误报率最低\n• 200条融合规则\n• 支持自定义组合', {
  x: 6.9, y: 3.1, w: 2.5, h: 0.85,
  fontSize: 9, color: COLORS.darkGray
});

addSlideNum(s8, slideNum, TOTAL);

// ========== SLIDE 9: 评分体系 ==========
slideNum++;
let s9 = pptx.addSlide();
s9.background = { color: COLORS.white };

s9.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.teal }
});

s9.addText('核心商业秘密风险管控评分体系', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s9.addText('两层决策：风险偏好确定上限，分值计算精细化执行动作', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

// 第一层：风险偏好
s9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 1.3, w: 4.4, h: 2.2,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.08
});

s9.addText('第一层：风险偏好', {
  x: 0.6, y: 1.4, w: 4.2, h: 0.35,
  fontSize: 13, color: COLORS.darkBg, bold: true
});
s9.addText('确定策略动作上限', {
  x: 0.6, y: 1.7, w: 4.2, h: 0.25,
  fontSize: 10, color: COLORS.midGray
});

const prefs = [
  { name: '激进', action: '阻断', color: COLORS.accent },
  { name: '平衡', action: '阻断', color: COLORS.warning },
  { name: '保守', action: '审批放行', color: COLORS.success },
];
prefs.forEach((p, i) => {
  const py = 2.0 + i * 0.5;
  s9.addShape(pptx.shapes.OVAL, {
    x: 0.7, y: py, w: 0.3, h: 0.3,
    fill: { color: p.color }
  });
  s9.addText(p.name, {
    x: 1.1, y: py, w: 1.0, h: 0.3,
    fontSize: 11, color: COLORS.darkBg, bold: true, valign: 'middle'
  });
  s9.addText('最高动作：' + p.action, {
    x: 2.1, y: py, w: 2.5, h: 0.3,
    fontSize: 10, color: COLORS.darkGray, valign: 'middle'
  });
});

// 第二层：分值计算
s9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.1, y: 1.3, w: 4.4, h: 2.2,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.08
});

s9.addText('第二层：分值计算', {
  x: 5.2, y: 1.4, w: 4.2, h: 0.35,
  fontSize: 13, color: COLORS.darkBg, bold: true
});
s9.addText('仅平衡偏好触发', {
  x: 5.2, y: 1.7, w: 4.2, h: 0.25,
  fontSize: 10, color: COLORS.midGray
});

// 公式
s9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.3, y: 2.0, w: 4.0, h: 0.45,
  fill: { color: COLORS.white },
  rectRadius: 0.05
});
s9.addText('分值 = 数据等级 × 行为风险 × 业务影响', {
  x: 5.3, y: 2.0, w: 4.0, h: 0.45,
  fontSize: 10, color: COLORS.darkBg, bold: true, align: 'center', valign: 'middle'
});

// 分值说明
const dimRows = [
  { dim: '数据等级', options: 'L1(1分) → L4(4分)' },
  { dim: '行为风险', options: '低(1分) → 中(2分) → 高(3分)' },
  { dim: '业务影响', options: '低(1分) → 中(2分) → 高(3分)' },
];
dimRows.forEach((r, i) => {
  s9.addText(r.dim + '：', {
    x: 5.3, y: 2.55 + i * 0.3, w: 1.0, h: 0.28,
    fontSize: 9, color: COLORS.darkGray, bold: true
  });
  s9.addText(r.options, {
    x: 6.3, y: 2.55 + i * 0.3, w: 3.0, h: 0.28,
    fontSize: 9, color: COLORS.darkGray
  });
});

// 阈值表
s9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 3.7, w: 9, h: 1.6,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.08
});

s9.addText('策略动作阈值（平衡偏好）', {
  x: 0.6, y: 3.8, w: 4, h: 0.35,
  fontSize: 13, color: COLORS.darkBg, bold: true
});

const thresholdHeaders = ['得分区间', '策略动作', '动作级别', '典型场景'];
const thW = [2.0, 2.0, 1.5, 3.0];
let thX = 0.6;
thresholdHeaders.forEach((h, i) => {
  s9.addShape(pptx.shapes.RECTANGLE, {
    x: thX, y: 4.2, w: thW[i], h: 0.3,
    fill: { color: COLORS.teal }
  });
  s9.addText(h, {
    x: thX, y: 4.2, w: thW[i], h: 0.3,
    fontSize: 9, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });
  thX += thW[i];
});

const thData = [
  { range: '27 ~ 36分', action: '阻断', level: '高', scene: 'L4数据 + 高风险 + 高影响' },
  { range: '9 ~ 26分', action: '审批放行', level: '中', scene: 'L3数据 + 中风险 + 中影响' },
  { range: '1 ~ 8分', action: '弹窗警告', level: '低', scene: 'L2数据 + 低风险 + 低影响' },
];
thData.forEach((r, ri) => {
  const ry = 4.5 + ri * 0.28;
  const vals = [r.range, r.action, r.level, r.scene];
  let tx2 = 0.6;
  vals.forEach((v, vi) => {
    s9.addShape(pptx.shapes.RECTANGLE, {
      x: tx2, y: ry, w: thW[vi], h: 0.26,
      fill: { color: ri % 2 === 0 ? COLORS.white : COLORS.lightGray },
      line: { color: COLORS.border, width: 0.3 }
    });
    s9.addText(v, {
      x: tx2 + 0.05, y: ry, w: thW[vi] - 0.1, h: 0.26,
      fontSize: 8, color: COLORS.darkGray, align: 'center', valign: 'middle'
    });
    tx2 += thW[vi];
  });
});

addSlideNum(s9, slideNum, TOTAL);

// ========== SLIDE 10: 建设阶段与策略动作 ==========
slideNum++;
let s10 = pptx.addSlide();
s10.background = { color: COLORS.white };

s10.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.orange }
});

s10.addText('数据安全建设三阶段', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s10.addText('阶段不同，策略不同。循序渐进，避免一步到位带来的业务冲击', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

const buildStages = [
  {
    stage: '前期', subtitle: '起步期',
    desc: '刚部署DLP/UEBA，\n需先建基线',
    duration: '约30天',
    action: '全部仅审计',
    focus: '积累行为数据\n建立基线\n发现高风险行为',
    color: COLORS.success,
    arrow: '→'
  },
  {
    stage: '中期', subtitle: '建设期',
    desc: '已建立基线，\n开始精细化运营',
    duration: '约3-6个月',
    action: 'L4弹窗警告\nL3以上审批',
    focus: '精细化阈值\n关键字扩充\n部门差异化',
    color: COLORS.warning,
    arrow: '→'
  },
  {
    stage: '后期', subtitle: '成熟期',
    desc: '运营机制完善，\n按规则严格执行',
    duration: '持续运营',
    action: '按风险偏好\n+分值阈值执行',
    focus: '融合策略\n精准阻断\n效率优化',
    color: COLORS.accent,
    arrow: ''
  },
];

buildStages.forEach((s, i) => {
  const x = 0.4 + i * 3.2;
  const y = 1.4;
  const w = 3.0;

  s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h: 3.9,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.1
  });

  // 顶部色条
  s10.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h: 0.15,
    fill: { color: s.color }
  });

  // 阶段
  s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.1, y: y + 0.25, w: 1.4, h: 0.35,
    fill: { color: s.color }, rectRadius: 0.05
  });
  s10.addText(s.stage, {
    x: x + 0.1, y: y + 0.25, w: 1.4, h: 0.35,
    fontSize: 13, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });

  s10.addText(s.subtitle, {
    x: x + 1.6, y: y + 0.25, w: 1.3, h: 0.35,
    fontSize: 12, color: COLORS.darkGray, valign: 'middle'
  });

  // 描述
  s10.addText(s.desc, {
    x: x + 0.1, y: y + 0.75, w: 2.8, h: 0.7,
    fontSize: 11, color: COLORS.darkGray
  });

  // 时间
  s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.1, y: y + 1.5, w: 2.8, h: 0.35,
    fill: { color: COLORS.white }, rectRadius: 0.05
  });
  s10.addText('周期：' + s.duration, {
    x: x + 0.1, y: y + 1.5, w: 2.8, h: 0.35,
    fontSize: 10, color: COLORS.darkGray, align: 'center', valign: 'middle'
  });

  // 策略动作
  s10.addText('策略动作', {
    x: x + 0.1, y: y + 2.0, w: 1.0, h: 0.25,
    fontSize: 9, color: COLORS.midGray, bold: true
  });
  s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.1, y: y + 2.25, w: 2.8, h: 0.45,
    fill: { color: s.color, transparency: 85 },
    rectRadius: 0.05
  });
  s10.addText(s.action, {
    x: x + 0.1, y: y + 2.25, w: 2.8, h: 0.45,
    fontSize: 10, color: COLORS.darkBg, align: 'center', valign: 'middle'
  });

  // 重点
  s10.addText('工作重点', {
    x: x + 0.1, y: y + 2.8, w: 1.0, h: 0.25,
    fontSize: 9, color: COLORS.midGray, bold: true
  });
  s10.addText(s.focus, {
    x: x + 0.1, y: y + 3.05, w: 2.8, h: 0.8,
    fontSize: 10, color: COLORS.darkGray
  });

  // 箭头
  if (s.arrow) {
    s10.addText(s.arrow, {
      x: x + w, y: y + 1.5, w: 0.2, h: 0.5,
      fontSize: 20, color: COLORS.midGray, align: 'center', valign: 'middle'
    });
  }
});

addSlideNum(s10, slideNum, TOTAL);

// ========== SLIDE 11: 法规依据 ==========
slideNum++;
let s11 = pptx.addSlide();
s11.background = { color: COLORS.white };

s11.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.darkGray }
});

s11.addText('策略设计的法规依据', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s11.addText('每条策略说明都应体现「依据《XX法》第XX条」，体现法有依据', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

const laws = [
  {
    name: '网络安全法',
    year: '2025修订',
    effective: '2026年1月1日',
    key: '第21条：等级保护\n第27条：禁止危害网络安全\n第31条：关键信息基础设施',
    color: COLORS.primary
  },
  {
    name: '数据安全法',
    year: '2021',
    effective: '2021年9月1日',
    key: '第21条：数据分类分级\n第27条：全流程安全管理制度\n第45条：违规处罚（最高50万）',
    color: COLORS.accent
  },
  {
    name: '个人信息保护法',
    year: '2021',
    effective: '2021年11月1日',
    key: '第28条：敏感个人信息定义\n第51条：加密/去标识化措施\n第66条：最高罚款营业额5%',
    color: COLORS.warning
  },
  {
    name: '网络数据安全管理条例',
    year: '2025',
    effective: '2025年1月1日',
    key: '第10条：技术安全措施\n第15条：数据安全负责人\n第27条：数据出境安全评估',
    color: COLORS.purple
  },
  {
    name: '关键信息基础设施保护条例',
    year: '2021',
    effective: '2021年9月1日',
    key: '第4条：重点保护范围\n第8条：设置安全管理机构\n第12条：安全检测评估',
    color: COLORS.teal
  },
];

laws.forEach((l, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.4 + col * 4.8;
  const y = 1.35 + row * 1.4;

  s11.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 4.6, h: 1.2,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.08
  });

  // 左侧色条
  s11.addShape(pptx.shapes.RECTANGLE, {
    x, y, w: 0.1, h: 1.2,
    fill: { color: l.color }
  });

  // 名称
  s11.addText(l.name, {
    x: x + 0.2, y: y + 0.08, w: 3.0, h: 0.3,
    fontSize: 12, color: COLORS.darkBg, bold: true
  });

  // 年份
  s11.addText(l.year + ' · ' + l.effective, {
    x: x + 3.2, y: y + 0.08, w: 1.2, h: 0.25,
    fontSize: 8, color: COLORS.midGray, align: 'right'
  });

  // 关键条款
  s11.addText(l.key, {
    x: x + 0.2, y: y + 0.4, w: 4.2, h: 0.75,
    fontSize: 9, color: COLORS.darkGray
  });
});

addSlideNum(s11, slideNum, TOTAL);

// ========== SLIDE 12: 工作流程 ==========
slideNum++;
let s12 = pptx.addSlide();
s12.background = { color: COLORS.white };

s12.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.purple }
});

s12.addText('5 步从需求到策略表', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

const steps = [
  { num: '1', title: '确认风险偏好', desc: '激进 / 平衡 / 保守', time: '1分钟' },
  { num: '2', title: '收集企业信息', desc: '行业 / 阶段 / 组织架构', time: '3分钟' },
  { num: '3', title: '精准化关键字', desc: '公司名+规章制度', time: '5分钟' },
  { num: '4', title: '生成策略表', desc: '三阶段策略设计', time: '2分钟' },
  { num: '5', title: '评分计算', desc: '分值+阈值+动作', time: '1分钟' },
];

steps.forEach((s, i) => {
  const x = 0.5 + i * 1.9;

  // 圆形序号
  s12.addShape(pptx.shapes.OVAL, {
    x: x + 0.55, y: 1.5, w: 0.7, h: 0.7,
    fill: { color: COLORS.purple }
  });
  s12.addText(s.num, {
    x: x + 0.55, y: 1.5, w: 0.7, h: 0.7,
    fontSize: 22, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });

  // 标题
  s12.addText(s.title, {
    x: x, y: 2.35, w: 1.8, h: 0.5,
    fontSize: 12, color: COLORS.darkBg, bold: true, align: 'center'
  });

  // 描述
  s12.addText(s.desc, {
    x: x, y: 2.85, w: 1.8, h: 0.5,
    fontSize: 10, color: COLORS.darkGray, align: 'center'
  });

  // 时间
  s12.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.2, y: 3.4, w: 1.4, h: 0.3,
    fill: { color: COLORS.lightGray }, rectRadius: 0.05
  });
  s12.addText(s.time, {
    x: x + 0.2, y: 3.4, w: 1.4, h: 0.3,
    fontSize: 9, color: COLORS.midGray, align: 'center', valign: 'middle'
  });

  // 箭头
  if (i < 4) {
    s12.addText('→', {
      x: x + 1.55, y: 1.55, w: 0.4, h: 0.6,
      fontSize: 18, color: COLORS.midGray, align: 'center', valign: 'middle'
    });
  }
});

// 底部说明
s12.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.0, w: 9, h: 1.2,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.1
});

s12.addText('使用 WorkBuddy，说出触发词即可开始', {
  x: 0.7, y: 4.1, w: 8.6, h: 0.35,
  fontSize: 13, color: COLORS.darkBg, bold: true
});

const triggers = [
  '"帮我设计研发部的数据安全防护策略"',
  '"提取财务部门的敏感数据关键字"',
  '"按保守风险偏好设计审批策略"',
  '"设计非工作时间+USB+核心商密的融合策略"',
];
triggers.forEach((t, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  s12.addText(t, {
    x: 0.7 + col * 4.5, y: 4.5 + row * 0.35, w: 4.3, h: 0.3,
    fontSize: 10, color: COLORS.primary
  });
});

addSlideNum(s12, slideNum, TOTAL);

// ========== SLIDE 13: 运营日报模块 ==========
slideNum++;
let s13 = pptx.addSlide();
s13.background = { color: COLORS.white };

s13.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.primary }
});

s13.addText('v2.0 新增：数据安全运营日报模块', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s13.addText('策略设计完成后，进入运营阶段，输出可读的运营报告评估策略效果', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

// 两种输入模式
const inputModes = [
  {
    mode: '方案A',
    name: '文字粘贴',
    subtitle: '轻量模式',
    icon: '📝',
    desc: '用户直接粘贴告警日志片段、聊天记录、部分表格数据',
    steps: ['读取文字内容', '识别关键字段', '风险分类', '输出日报'],
    color: COLORS.success
  },
  {
    mode: '方案B',
    name: '文件上传',
    subtitle: '完整模式',
    icon: '📁',
    desc: '上传原始告警日志文件（CSV/Excel/JSON/TXT）',
    steps: ['读取文件', '自动识别格式', '字段映射', '生成报告'],
    color: COLORS.primary
  },
];

inputModes.forEach((m, i) => {
  const x = 0.5 + i * 4.7;
  const y = 1.4;

  s13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 4.5, h: 2.5,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.1
  });

  // 顶部色条
  s13.addShape(pptx.shapes.RECTANGLE, {
    x, y, w: 4.5, h: 0.12,
    fill: { color: m.color }
  });

  // 模式标签
  s13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.15, y: y + 0.25, w: 0.8, h: 0.3,
    fill: { color: m.color }, rectRadius: 0.05
  });
  s13.addText(m.mode, {
    x: x + 0.15, y: y + 0.25, w: 0.8, h: 0.3,
    fontSize: 10, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });

  s13.addText(m.name, {
    x: x + 1.05, y: y + 0.22, w: 1.5, h: 0.35,
    fontSize: 14, color: COLORS.darkBg, bold: true
  });

  s13.addText(m.subtitle, {
    x: x + 2.5, y: y + 0.25, w: 1.8, h: 0.3,
    fontSize: 10, color: COLORS.midGray, align: 'right'
  });

  // 描述
  s13.addText(m.desc, {
    x: x + 0.15, y: y + 0.65, w: 4.2, h: 0.5,
    fontSize: 10, color: COLORS.darkGray
  });

  // 流程步骤
  m.steps.forEach((step, si) => {
    const sx = x + 0.15 + si * 1.0;
    s13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: sx, y: y + 1.3, w: 0.9, h: 0.5,
      fill: { color: COLORS.white }, rectRadius: 0.05
    });
    s13.addText(step, {
      x: sx, y: y + 1.3, w: 0.9, h: 0.5,
      fontSize: 7, color: COLORS.darkGray, align: 'center', valign: 'middle'
    });
    if (si < 3) {
      s13.addText('→', {
        x: sx + 0.85, y: y + 1.35, w: 0.2, h: 0.4,
        fontSize: 10, color: COLORS.midGray
      });
    }
  });

  // 触发词
  s13.addText('触发词：' + (i === 0 ? '"帮我整理日报"' : '"上传日志文件"'), {
    x: x + 0.15, y: y + 1.95, w: 4.2, h: 0.3,
    fontSize: 9, color: COLORS.primary
  });
});

// 日报结构
s13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.05, w: 9, h: 1.2,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.1
});

s13.addText('运营日报结构（7个部分）', {
  x: 0.7, y: 4.15, w: 4, h: 0.3,
  fontSize: 12, color: COLORS.darkBg, bold: true
});

const reportParts = ['防护范围概览', '整体运营情况', '重点风险事件', '部门分布分析', '趋势对比', '处置建议', '附录'];
reportParts.forEach((p, i) => {
  const px = 0.7 + i * 1.25;
  s13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: px, y: 4.55, w: 1.15, h: 0.55,
    fill: { color: COLORS.white }, rectRadius: 0.05
  });
  s13.addText(p, {
    x: px, y: 4.55, w: 1.15, h: 0.55,
    fontSize: 8, color: COLORS.darkGray, align: 'center', valign: 'middle'
  });
});

addSlideNum(s13, slideNum, TOTAL);

// ========== SLIDE 14: 适用产品 ==========
slideNum++;
let s14 = pptx.addSlide();
s14.background = { color: COLORS.white };

s14.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.teal }
});

s14.addText('适用产品与数据格式', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s14.addText('支持多种 DLP/UEBA 产品的告警日志解析', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

// 产品列表
const products = [
  { name: '奇安信 DLP', type: '国内头部' },
  { name: '安天 DLP', type: '国内头部' },
  { name: '天空卫士 DLP', type: '国内头部' },
  { name: 'Symantec DLP', type: '国际厂商' },
  { name: '微软 Purview DLP', type: '国际厂商' },
  { name: '定制化/自研平台', type: '其他' },
];

products.forEach((p, i) => {
  const col = i % 3;
  const row = Math.floor(i / 3);
  const x = 0.5 + col * 3.1;
  const y = 1.5 + row * 0.9;

  s14.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 2.9, h: 0.7,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.08
  });

  s14.addText(p.name, {
    x: x + 0.15, y: y + 0.1, w: 2.0, h: 0.3,
    fontSize: 13, color: COLORS.darkBg, bold: true
  });

  s14.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 2.15, y: y + 0.12, w: 0.6, h: 0.25,
    fill: { color: p.type === '国内头部' ? COLORS.accent : (p.type === '国际厂商' ? COLORS.primary : COLORS.teal) },
    rectRadius: 0.03
  });
  s14.addText(p.type, {
    x: x + 2.15, y: y + 0.12, w: 0.6, h: 0.25,
    fontSize: 6, color: COLORS.white, align: 'center', valign: 'middle'
  });

  s14.addText('支持CSV/Excel/JSON/TXT格式导入', {
    x: x + 0.15, y: y + 0.4, w: 2.6, h: 0.25,
    fontSize: 9, color: COLORS.midGray
  });
});

// 标准字段体系
s14.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 3.5, w: 9, h: 1.8,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.1
});

s14.addText('统一映射到 11 个标准字段', {
  x: 0.7, y: 3.6, w: 4, h: 0.35,
  fontSize: 13, color: COLORS.darkBg, bold: true
});

const fields = [
  '用户/发件人', '时间/发件时间', '外发通道', '命中策略名称',
  '命中关键词', '文件名/邮件附件', '目标邮箱', '风险等级',
  '策略分值', '建议动作', '疑似违规详情'
];

fields.forEach((f, i) => {
  const col = i % 4;
  const row = Math.floor(i / 4);
  const fx = 0.7 + col * 2.2;
  const fy = 4.05 + row * 0.55;

  s14.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: fx, y: fy, w: 2.0, h: 0.4,
    fill: { color: COLORS.white },
    rectRadius: 0.05
  });

  const num = String(i + 1).padStart(2, '0');
  s14.addShape(pptx.shapes.OVAL, {
    x: fx + 0.05, y: fy + 0.05, w: 0.3, h: 0.3,
    fill: { color: COLORS.teal }
  });
  s14.addText(num, {
    x: fx + 0.05, y: fy + 0.05, w: 0.3, h: 0.3,
    fontSize: 8, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });

  s14.addText(f, {
    x: fx + 0.4, y: fy, w: 1.55, h: 0.4,
    fontSize: 9, color: COLORS.darkGray, valign: 'middle'
  });
});

addSlideNum(s14, slideNum, TOTAL);

// ========== SLIDE 15: 使用效果 ==========
slideNum++;
let s15 = pptx.addSlide();
s15.background = { color: COLORS.white };

s15.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.accent }
});

s15.addText('面向不同人群的使用效果', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

const audiences = [
  {
    role: '安全管理员',
    pain: '策略配置复杂，不知从何下手？331条策略+评分体系，直接套用',
    value: '✓ 5分钟生成部门策略表\n✓ 法规依据自动标注\n✓ 分阶段推进不迷茫',
    color: COLORS.accent
  },
  {
    role: 'IT 负责人',
    pain: '部门需求多，沟通成本高？输出标准化策略表，直接评审',
    value: '✓ 结构化策略输出\n✓ 两段式说明（选择依据+设置依据）\n✓ 支持飞书/Excel多格式',
    color: COLORS.primary
  },
  {
    role: '合规负责人',
    pain: '监管检查如何证明策略有效性？日报模块量化运营成果',
    value: '✓ 运营日报自动生成\n✓ 处置建议直接落地\n✓ 版本更新记录可追溯',
    color: COLORS.success
  },
  {
    role: '安全厂商售前',
    pain: 'POC方案如何快速产出差异化价值？方法论+工具双重背书',
    value: '✓ 《以人为本》方法论支撑\n✓ 431条策略库快速配置\n✓ 评分体系增强专业感',
    color: COLORS.purple
  },
];

audiences.forEach((a, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.4 + col * 4.8;
  const y = 1.4 + row * 2.0;

  s15.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 4.6, h: 1.8,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.1
  });

  // 左侧色条
  s15.addShape(pptx.shapes.RECTANGLE, {
    x, y, w: 0.12, h: 1.8,
    fill: { color: a.color }
  });

  // 角色
  s15.addText(a.role, {
    x: x + 0.25, y: y + 0.1, w: 4.2, h: 0.35,
    fontSize: 14, color: COLORS.darkBg, bold: true
  });

  // 痛点
  s15.addText('痛点：' + a.pain, {
    x: x + 0.25, y: y + 0.5, w: 4.2, h: 0.5,
    fontSize: 10, color: COLORS.midGray
  });

  // 价值
  s15.addText(a.value, {
    x: x + 0.25, y: y + 1.0, w: 4.2, h: 0.75,
    fontSize: 10, color: a.color
  });
});

addSlideNum(s15, slideNum, TOTAL);

// ========== SLIDE 16: 文件结构 ==========
slideNum++;
let s16 = pptx.addSlide();
s16.background = { color: COLORS.white };

s16.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.darkGray }
});

s16.addText('项目文件结构', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

const fileStruct = `data-security-policy-generator/
├── SKILL.md                              # 技能主文件（AI 读取的指令）
├── README.md                             # 产品说明文档
├── references/                           # 参考资料目录
│   ├── 01_strategy_catalog.md           # 策略目录示例（预览）
│   ├── 02_department_mapping.md         # 部门×策略适用性矩阵
│   ├── 03_legal_framework.md            # 数据安全法规手册
│   ├── 04_bu_keywords.md                # BU关键字字典示例（预览）
│   ├── 05_keyword_taxonomy.md           # 扩充字典示例（预览）
│   └── api_reference.md                  # DLP/UEBA API 对接参考
├── SOP/                                 # 标准操作流程
│   ├── 01_mvp_strategy.md              # MVP 策略设计流程
│   ├── 02_business_data_extract.md     # 业务数据关键字提取
│   ├── 03_fusion_strategy.md           # 融合策略设计
│   └── 04_scoring_system.md            # 策略评分体系
└── scripts/
    └── __pycache__/
        └── strategy_lib.cpython-313.pyc  # 🔒 加密策略库（431条策略）`;

s16.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 1.3, w: 9, h: 3.9,
  fill: { color: '1E293B' },
  rectRadius: 0.1
});

s16.addText(fileStruct, {
  x: 0.7, y: 1.4, w: 8.6, h: 3.7,
  fontSize: 9, color: '94A3B8', fontFace: 'Courier New'
});

addSlideNum(s16, slideNum, TOTAL);

// ========== SLIDE 17: 版本更新记录 ==========
slideNum++;
let s17 = pptx.addSlide();
s17.background = { color: COLORS.white };

s17.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.warning }
});

s17.addText('版本更新记录', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

const versions = [
  { ver: 'v2.1', date: '2026-04-23', content: '新增禁止原始数据查询条款，保护人员隐私信息', type: '小功能' },
  { ver: 'v2.0', date: '2026-04-23', content: '大功能升级：新增运营日报模块（方案A文字/方案B文件）；新增报告脱敏规则；新增版本规范', type: '大功能' },
  { ver: 'v1.3', date: '2026-04-23', content: '修正报告脱敏规则，区分公司名称（可脱敏）与人员行为信息（需保留）', type: '优化' },
  { ver: 'v1.2', date: '2026-04-23', content: '新增行为基线四原则（数据透明/同事语气/时间核查/错误响应）；新增策略评分体系', type: '新增' },
  { ver: 'v1.1', date: '2026-04-23', content: '内置策略库加密处理；新增示例预览文件；触发词优化', type: '优化' },
  { ver: 'v1.0', date: '2026-04-23', content: '初始版本：三阶段策略设计框架；331条行为管控策略；200条融合策略；13个部门关键字映射', type: '初始' },
];

const verHeaders = ['版本', '日期', '变更内容', '类型'];
const verW = [0.8, 1.1, 5.5, 1.2];
let vx = 0.5;
verHeaders.forEach((h, i) => {
  s17.addShape(pptx.shapes.RECTANGLE, {
    x: vx, y: 1.3, w: verW[i], h: 0.35,
    fill: { color: COLORS.warning }
  });
  s17.addText(h, {
    x: vx, y: 1.3, w: verW[i], h: 0.35,
    fontSize: 10, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });
  vx += verW[i];
});

versions.forEach((v, ri) => {
  const ry = 1.65 + ri * 0.6;
  const bgColor = ri === 0 ? 'FFF7ED' : (ri % 2 === 0 ? COLORS.lightGray : COLORS.white);
  const vals = [v.ver, v.date, v.content, v.type];
  const typeColor = {
    '大功能': COLORS.accent, '小功能': COLORS.primary,
    '优化': COLORS.warning, '新增': COLORS.purple,
    '初始': COLORS.success
  };

  let cx = 0.5;
  vals.forEach((val, vi) => {
    s17.addShape(pptx.shapes.RECTANGLE, {
      x: cx, y: ry, w: verW[vi], h: 0.55,
      fill: { color: bgColor },
      line: { color: COLORS.border, width: 0.5 }
    });

    const fontColor = vi === 0 ? COLORS.warning : (vi === 3 ? typeColor[v.type] : COLORS.darkGray);
    const bold = vi === 0;
    s17.addText(val, {
      x: cx + 0.05, y: ry, w: verW[vi] - 0.1, h: 0.55,
      fontSize: 9, color: fontColor, bold, align: vi === 0 || vi === 1 || vi === 3 ? 'center' : 'left', valign: 'middle'
    });
    cx += verW[vi];
  });
});

// 版本规范说明
s17.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 5.3, w: 9, h: 0.0,
  fill: { color: COLORS.lightGray },
  rectRadius: 0
});

addSlideNum(s17, slideNum, TOTAL);

// ========== SLIDE 18: 术语表 ==========
slideNum++;
let s18 = pptx.addSlide();
s18.background = { color: COLORS.white };

s18.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.purple }
});

s18.addText('核心术语速查表', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

const terms = [
  { term: 'DLP', full: 'Data Loss Prevention', meaning: '数据防泄漏，通过策略检测和阻止敏感数据外发' },
  { term: 'UEBA', full: 'User and Entity Behavior Analytics', meaning: '用户和实体行为分析，通过行为异常检测发现风险' },
  { term: 'L4', full: '核心商密', meaning: '泄露后会对企业造成严重损害的数据，如战略规划、高管薪酬' },
  { term: 'L3', full: '普通商密', meaning: '具有商业价值但泄露影响有限的数据，如客户名单' },
  { term: 'L2', full: '内部知悉', meaning: '仅限内部使用，泄露后影响可控的数据，如项目进度' },
  { term: 'L1', full: '公开信息', meaning: '可对外公开的数据，无需特别管控' },
  { term: 'R策略', full: '规则型策略', meaning: '基于时间/频率/通道等条件触发的策略' },
  { term: 'C策略', full: '内容型策略', meaning: '基于正则表达式匹配敏感内容的策略' },
  { term: 'S策略', full: '敏感行为型策略', meaning: '检测规避管控的行为，如压缩、伪装、隐藏' },
  { term: '融合策略', full: '行为×内容融合', meaning: '行为和内容都为敏感时才触发，精准度最高' },
];

terms.forEach((t, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.4 + col * 4.8;
  const y = 1.35 + row * 0.8;

  s18.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 4.6, h: 0.7,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.06
  });

  // 术语标签
  s18.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.1, y: y + 0.1, w: 0.8, h: 0.5,
    fill: { color: COLORS.purple },
    rectRadius: 0.05
  });
  s18.addText(t.term, {
    x: x + 0.1, y: y + 0.1, w: 0.8, h: 0.5,
    fontSize: 10, color: COLORS.white, bold: true, align: 'center', valign: 'middle'
  });

  // 全称
  s18.addText(t.full, {
    x: x + 1.0, y: y + 0.08, w: 3.4, h: 0.28,
    fontSize: 9, color: COLORS.darkGray
  });

  // 含义
  s18.addText(t.meaning, {
    x: x + 1.0, y: y + 0.38, w: 3.4, h: 0.28,
    fontSize: 9, color: COLORS.midGray
  });
});

addSlideNum(s18, slideNum, TOTAL);

// ========== SLIDE 19: 沟通风格适配 ==========
slideNum++;
let s19 = pptx.addSlide();
s19.background = { color: COLORS.white };

s19.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.success }
});

s19.addText('沟通风格适配', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

s19.addText('Skill 被调用时，第一步先了解用户的沟通偏好', {
  x: 0.5, y: 0.85, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray
});

// 技术模式
s19.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 1.4, w: 4.4, h: 3.3,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.1
});

s19.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.4, w: 4.4, h: 0.15,
  fill: { color: COLORS.primary }
});

s19.addText('专业模式', {
  x: 0.7, y: 1.65, w: 4.0, h: 0.4,
  fontSize: 16, color: COLORS.darkBg, bold: true
});

s19.addText('适合：IT/安全/运维人员', {
  x: 0.7, y: 2.05, w: 4.0, h: 0.3,
  fontSize: 11, color: COLORS.primary
});

const techFeatures = [
  '使用专业术语：DLP、UEBA、分值阈值',
  '输出技术参数表、配置示例',
  '默认用户已了解 L4/L3/L2 含义',
  '引用法规条款编号',
  '讨论分值公式、评分体系',
];

s19.addText(techFeatures.map((f, i) => ({
  text: '✓ ' + f,
  options: { breakLine: i < techFeatures.length - 1 }
})), {
  x: 0.7, y: 2.45, w: 4.0, h: 2.1,
  fontSize: 11, color: COLORS.darkGray,
  paraSpaceAfter: 8
});

// 业务模式
s19.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.1, y: 1.4, w: 4.4, h: 3.3,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.1
});

s19.addShape(pptx.shapes.RECTANGLE, {
  x: 5.1, y: 1.4, w: 4.4, h: 0.15,
  fill: { color: COLORS.success }
});

s19.addText('大白话模式', {
  x: 5.3, y: 1.65, w: 4.0, h: 0.4,
  fontSize: 16, color: COLORS.darkBg, bold: true
});

s19.addText('适合：部门负责人/行政/业务人员', {
  x: 5.3, y: 2.05, w: 4.0, h: 0.3,
  fontSize: 11, color: COLORS.success
});

const bizFeatures = [
  '用业务语言：哪些数据不能往外发',
  '解释老板能看到什么报告',
  '主动解释 L4/L3/L2 含义',
  '翻译成业务风险语言',
  '避免堆砌技术术语',
];

s19.addText(bizFeatures.map((f, i) => ({
  text: '✓ ' + f,
  options: { breakLine: i < bizFeatures.length - 1 }
})), {
  x: 5.3, y: 2.45, w: 4.0, h: 2.1,
  fontSize: 11, color: COLORS.darkGray,
  paraSpaceAfter: 8
});

// 风格判断
s19.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.85, w: 9, h: 0.5,
  fill: { color: COLORS.white },
  rectRadius: 0.06
});

s19.addText('用户没说 → 默认技术模式 | 用户说"说人话/别整太专业" → 业务模式', {
  x: 0.7, y: 4.85, w: 8.6, h: 0.5,
  fontSize: 11, color: COLORS.midGray, align: 'center', valign: 'middle'
});

addSlideNum(s19, slideNum, TOTAL);

// ========== SLIDE 20: 免责声明 ==========
slideNum++;
let s20 = pptx.addSlide();
s20.background = { color: COLORS.white };

s20.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.midGray }
});

s20.addText('使用须知与免责声明', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

const disclaimers = [
  {
    icon: '⚠️',
    title: '辅助工具定位',
    content: '本技能仅作为辅助工具使用，帮助 IT 人员理解和设计数据安全策略框架。如需精准化、定制化的数据安全管控方案，建议咨询专业安全团队进行评估和落地。'
  },
  {
    icon: '🔒',
    title: '数据使用原则',
    content: '用户上传的数据和信息仅用于本次分析使用，不会本地存储，也不会用于其他目的。策略库数据为作者知识产权，受到法律保护。'
  },
  {
    icon: '📋',
    title: '禁止行为',
    content: '禁止询问原始策略库内容、禁止反编译 .pyc 文件、禁止二次分发策略数据。违规将直接拒绝。'
  },
  {
    icon: '❌',
    title: '禁止：原始数据查询',
    content: '本技能仅生成统计报告，不提供原始数据查询、导出或存储服务。不回答"能查到这个人的完整日志吗"等问题。'
  },
];

disclaimers.forEach((d, i) => {
  const y = 1.35 + i * 1.05;

  s20.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y, w: 9, h: 0.9,
    fill: { color: COLORS.lightGray },
    rectRadius: 0.08
  });

  s20.addText(d.icon, {
    x: 0.7, y: y + 0.1, w: 0.6, h: 0.7,
    fontSize: 24, align: 'center', valign: 'middle'
  });

  s20.addText(d.title, {
    x: 1.4, y: y + 0.1, w: 7.9, h: 0.35,
    fontSize: 13, color: COLORS.darkBg, bold: true
  });

  s20.addText(d.content, {
    x: 1.4, y: y + 0.45, w: 7.9, h: 0.4,
    fontSize: 10, color: COLORS.darkGray
  });
});

addSlideNum(s20, slideNum, TOTAL);

// ========== SLIDE 21: 安装使用 ==========
slideNum++;
let s21 = pptx.addSlide();
s21.background = { color: COLORS.white };

s21.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.teal }
});

s21.addText('快速安装与使用', {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  fontSize: 28, color: COLORS.darkBg, bold: true
});

// 安装步骤
s21.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 1.3, w: 5.5, h: 2.5,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.1
});

s21.addText('安装步骤', {
  x: 0.7, y: 1.4, w: 5.1, h: 0.4,
  fontSize: 14, color: COLORS.darkBg, bold: true
});

const installSteps = [
  '1. 下载 data-security-policy-generator-v2.0-release.zip',
  '2. 解压到 WorkBuddy Skills 目录：',
  '   unzip data-security-policy-generator-v2.0-release.zip \\\\',
  '     -d ~/.workbuddy/skills/data-security-policy-generator/',
  '3. 在 WorkBuddy 中说出触发词即可开始使用',
];

s21.addText(installSteps.join('\n'), {
  x: 0.7, y: 1.85, w: 5.1, h: 1.8,
  fontSize: 10, color: COLORS.darkGray, fontFace: 'Courier New'
});

// 触发词示例
s21.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 6.2, y: 1.3, w: 3.3, h: 2.5,
  fill: { color: COLORS.lightGray },
  rectRadius: 0.1
});

s21.addText('常用触发词', {
  x: 6.4, y: 1.4, w: 2.9, h: 0.4,
  fontSize: 14, color: COLORS.darkBg, bold: true
});

const triggerWords = [
  '"帮我设计研发部的策略"',
  '"提取财务部门关键字"',
  '"按保守偏好设计审批"',
  '"生成数据安全运营日报"',
  '"帮我分析日志输出报告"',
];

s21.addText(triggerWords.map((t, i) => ({
  text: t,
  options: { breakLine: i < triggerWords.length - 1 }
})), {
  x: 6.4, y: 1.85, w: 2.9, h: 1.8,
  fontSize: 9, color: COLORS.primary,
  paraSpaceAfter: 8
});

// GitHub 信息
s21.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.0, w: 9, h: 1.2,
  fill: { color: '1E293B' },
  rectRadius: 0.1
});

s21.addText('GitHub 仓库', {
  x: 0.7, y: 4.1, w: 8.6, h: 0.35,
  fontSize: 13, color: COLORS.white, bold: true
});

s21.addText('github.com/LiuKai-DS/data-security-policy-generator', {
  x: 0.7, y: 4.45, w: 8.6, h: 0.3,
  fontSize: 14, color: COLORS.primary
});

s21.addText('MIT License · 欢迎 Star ⭐ 和 Fork · 持续更新中', {
  x: 0.7, y: 4.8, w: 8.6, h: 0.3,
  fontSize: 11, color: COLORS.midGray
});

addSlideNum(s21, slideNum, TOTAL);

// ========== SLIDE 22: 结束页 ==========
slideNum++;
let s22 = pptx.addSlide();
s22.background = { color: COLORS.darkBg };

// 装饰
s22.addShape(pptx.shapes.OVAL, {
  x: -1.5, y: -1.5, w: 4, h: 4,
  fill: { color: COLORS.primary, transparency: 85 }
});
s22.addShape(pptx.shapes.OVAL, {
  x: 8, y: 3.5, w: 3, h: 3,
  fill: { color: COLORS.accent, transparency: 85 }
});

s22.addText('以人为本的数据安全运营', {
  x: 0.5, y: 1.8, w: 9, h: 0.8,
  fontSize: 36, color: COLORS.white, bold: true, align: 'center'
});

s22.addText('策略是死的，人是活的', {
  x: 0.5, y: 2.7, w: 9, h: 0.5,
  fontSize: 20, color: COLORS.primary, align: 'center'
});

// 分隔线
s22.addShape(pptx.shapes.RECTANGLE, {
  x: 3.5, y: 3.4, w: 3, h: 0.02,
  fill: { color: COLORS.midGray }
});

s22.addText('github.com/LiuKai-DS/data-security-policy-generator', {
  x: 0.5, y: 3.7, w: 9, h: 0.4,
  fontSize: 14, color: COLORS.midGray, align: 'center'
});

s22.addText('MIT License · 欢迎 Star ⭐', {
  x: 0.5, y: 4.2, w: 9, h: 0.3,
  fontSize: 12, color: COLORS.midGray, align: 'center'
});

s22.addText('联系方式：liukai31415926@163.com', {
  x: 0.5, y: 4.6, w: 9, h: 0.3,
  fontSize: 11, color: COLORS.midGray, align: 'center'
});

// ========== 保存 ==========
const outPath = '/Users/law42pulse/WorkBuddy/20260410110313/data-security-policy-generator-v2.0-release/data-security-policy-generator-产品介绍v2.0.pptx';
pptx.writeFile({ fileName: outPath })
  .then(() => {
    console.log('PPT 生成成功：' + outPath);
    console.log('共 ' + slideNum + ' 页');
  })
  .catch(err => console.error('生成失败:', err));
