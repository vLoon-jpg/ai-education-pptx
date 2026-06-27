// hermes-race.js — Hermes builds his own PPTX while Claude races
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Hermes";
pres.title = "How AI is Changing High School Education";

const C = {
  dark:     "0F0722",
  purple:   "5B21B6",
  indigo:   "4338CA",
  gold:     "F59E0B",
  lightBg:  "F8FAFC",
  cardBg:   "F1F5F9",
  textDark: "1E293B",
  textMid:  "64748B",
  white:    "FFFFFF",
};

const makeShadow = () => ({ type:"outer", blur:4, offset:1, angle:135, color:"000000", opacity:0.08 });

// SLIDE 1 — Title
(() => {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addText("How AI is Changing\nHigh School Education", {
    x:0.8, y:1.2, w:8.4, h:1.6, fontSize:40, fontFace:"Arial Black", color:C.white, bold:true, align:"center", margin:0,
  });
  s.addShape(pres.shapes.RECTANGLE, { x:3.5, y:3.0, w:3.0, h:0.04, fill:{color:C.gold} });
  s.addText("Levy • Member 2 • Member 3 • Member 4 • Member 5", {
    x:1, y:3.4, w:8, h:0.5, fontSize:13, fontFace:"Calibri", color:C.textMid, align:"center", margin:0,
  });
})();

// SLIDE 2 — What is AI
(() => {
  const s = pres.addSlide();
  s.background = { color: C.lightBg };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.85, w:0.08, h:2.5, fill:{color:C.purple} });
  s.addText("What Is AI?", { x:0.8, y:0.3, w:8, h:0.6, fontSize:32, fontFace:"Arial Black", color:C.textDark, bold:true, margin:0 });

  s.addText([
    { text:"AI = Artificial Intelligence", options:{bold:true, fontSize:18, color:C.indigo, breakLine:true} },
    { text:"", options:{breakLine:true, fontSize:8} },
    { text:"• ", options:{breakLine:true} },
    { text:"Machines that can learn, reason, and make decisions", options:{breakLine:true, fontSize:14, color:C.textMid} },
    { text:"• Chatbots like ChatGPT think like a super-fast research assistant", options:{breakLine:true, fontSize:14, color:C.textMid} },
    { text:"• AI already in your pocket: Siri, Google Translate, TikTok recommendations", options:{breakLine:true, fontSize:14, color:C.textMid} },
    { text:"• Not magic — just math, data, and pattern recognition at huge scale", options:{fontSize:14, color:C.textMid} },
  ], { x:0.8, y:1.6, w:8.4, h:3.0, fontFace:"Calibri", valign:"top", margin:0 });

  s.addShape(pres.shapes.RECTANGLE, { x:6.5, y:4.2, w:3.0, h:0.9, fill:{color:C.indigo} });
  s.addText("Fun Fact\nAI has been around since the 1950s!\nIt just got really good in the last 5 years.", {
    x:6.65, y:4.25, w:2.7, h:0.8, fontSize:9, fontFace:"Calibri", color:C.white, align:"center", valign:"middle", margin:0,
  });
})();

// SLIDE 3 — AI in Learning
(() => {
  const s = pres.addSlide();
  s.background = { color: C.lightBg };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.85, w:0.08, h:2.5, fill:{color:C.purple} });
  s.addText("AI in Learning", { x:0.8, y:0.3, w:8, h:0.6, fontSize:32, fontFace:"Arial Black", color:C.textDark, bold:true, margin:0 });

  const cards = [
    { emoji:"🎯", title:"Personalized\nTutoring", desc:"AI adapts to YOUR\npace and style —\nno two students\nget the same lesson" },
    { emoji:"⚡", title:"Instant\nFeedback", desc:"Essays graded in\nseconds, mistakes\nexplained immediately\n— no more waiting\na week" },
    { emoji:"🌍", title:"Language\nLearning", desc:"Real-time translation\nand pronunciation\ncoaching — like having\na native speaker\n24/7" },
  ];

  cards.forEach((c, i) => {
    const x = 0.8 + i*3.0;
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.6, w:2.7, h:3.0, fill:{color:C.cardBg}, shadow:makeShadow() });
    s.addText(c.emoji, { x, y:1.75, w:2.7, h:0.5, fontSize:28, align:"center", margin:0 });
    s.addText(c.title, { x:x+0.15, y:2.35, w:2.4, h:0.7, fontSize:13, fontFace:"Arial Black", color:C.indigo, bold:true, margin:0 });
    s.addText(c.desc, { x:x+0.15, y:3.15, w:2.4, h:1.2, fontSize:11, fontFace:"Calibri", color:C.textMid, margin:0 });
  });
})();

// SLIDE 4 — AI Tools
(() => {
  const s = pres.addSlide();
  s.background = { color: C.lightBg };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.85, w:0.08, h:2.5, fill:{color:C.purple} });
  s.addText("AI Tools Students Use", { x:0.8, y:0.3, w:8, h:0.6, fontSize:32, fontFace:"Arial Black", color:C.textDark, bold:true, margin:0 });

  const tools = [
    ["ChatGPT", "Writing help, research, brainstorming, explaining concepts"],
    ["Grammarly", "Grammar & tone checking, writing suggestions"],
    ["Photomath", "Scan a math problem → step-by-step solution"],
    ["Quizlet AI", "Auto-generates flashcards & practice tests"],
  ];

  tools.forEach((t, i) => {
    const y = 1.6 + i*0.85;
    s.addShape(pres.shapes.RECTANGLE, { x:0.8, y, w:8.4, h:0.65, fill:{color:C.cardBg}, shadow:makeShadow() });
    s.addText(t[0], { x:1.0, y:y+0.08, w:2.0, h:0.5, fontSize:14, fontFace:"Arial Black", color:C.indigo, bold:true, margin:0 });
    s.addText(t[1], { x:3.2, y:y+0.08, w:5.8, h:0.5, fontSize:12, fontFace:"Calibri", color:C.textMid, valign:"middle", margin:0 });
  });

  s.addText("⚠️  Always check school AI policies before using these tools for assignments", {
    x:0.8, y:5.1, w:8.4, h:0.3, fontSize:10, fontFace:"Calibri", color:C.textMid, italic:true, margin:0,
  });
})();

// SLIDE 5 — Benefits & Concerns
(() => {
  const s = pres.addSlide();
  s.background = { color: C.lightBg };
  s.addText("Benefits & Concerns", { x:0.8, y:0.3, w:8, h:0.6, fontSize:32, fontFace:"Arial Black", color:C.textDark, bold:true, margin:0 });

  // Benefits
  s.addShape(pres.shapes.RECTANGLE, { x:0.8, y:1.3, w:4.0, h:3.5, fill:{color:C.cardBg} });
  s.addText("✅ Benefits", { x:1.0, y:1.4, w:3.6, h:0.4, fontSize:16, fontFace:"Arial Black", color:C.indigo, bold:true, margin:0 });
  s.addText([
    { text:"Learn at your own speed", options:{bullet:true, breakLine:true, fontSize:13, color:C.textMid} },
    { text:"24/7 homework help available", options:{bullet:true, breakLine:true, fontSize:13, color:C.textMid} },
    { text:"Prepares you for an AI-powered future workforce", options:{bullet:true, breakLine:true, fontSize:13, color:C.textMid} },
    { text:"Makes education more accessible globally", options:{bullet:true, fontSize:13, color:C.textMid} },
  ], { x:1.0, y:2.0, w:3.6, h:2.5, fontFace:"Calibri", valign:"top", margin:0 });

  // Concerns
  s.addShape(pres.shapes.RECTANGLE, { x:5.2, y:1.3, w:4.0, h:3.5, fill:{color:C.cardBg} });
  s.addText("⚠️ Concerns", { x:5.4, y:1.4, w:3.6, h:0.4, fontSize:16, fontFace:"Arial Black", color:"DC2626", bold:true, margin:0 });
  s.addText([
    { text:"Over-reliance — losing critical thinking", options:{bullet:true, breakLine:true, fontSize:13, color:C.textMid} },
    { text:"Plagiarism & academic dishonesty risks", options:{bullet:true, breakLine:true, fontSize:13, color:C.textMid} },
    { text:"Not all students have equal access (digital divide)", options:{bullet:true, breakLine:true, fontSize:13, color:C.textMid} },
    { text:"AI can be confidently wrong (hallucinations)", options:{bullet:true, fontSize:13, color:C.textMid} },
  ], { x:5.4, y:2.0, w:3.6, h:2.5, fontFace:"Calibri", valign:"top", margin:0 });
})();

// SLIDE 6 — Takeaways + Q&A
(() => {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addText("Key Takeaways", { x:1, y:0.5, w:8, h:0.7, fontSize:34, fontFace:"Arial Black", color:C.white, bold:true, margin:0 });

  const pts = [
    "AI is already in your daily life — from TikTok to Google Translate",
    "It can personalize learning, give instant feedback, and help with languages",
    "Tools like ChatGPT and Grammarly are useful but need responsible use",
    "The goal isn't replacing teachers — it's giving them superpowers",
  ];
  pts.forEach((pt, i) => {
    const y = 1.5 + i*0.85;
    s.addShape(pres.shapes.OVAL, { x:1.0, y:y+0.05, w:0.4, h:0.4, fill:{color:C.gold} });
    s.addText(String(i+1), { x:1.0, y:y+0.05, w:0.4, h:0.4, fontSize:16, fontFace:"Arial Black", color:C.dark, align:"center", valign:"middle", margin:0 });
    s.addText(pt, { x:1.7, y:y+0.05, w:7.0, h:0.4, fontSize:13, fontFace:"Calibri", color:C.white, valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:4, y:4.6, w:2, h:0.03, fill:{color:C.gold} });
  s.addText("Questions?  •  Thank you for listening!", { x:1, y:4.85, w:8, h:0.3, fontSize:12, fontFace:"Calibri", color:C.textMid, align:"center", margin:0 });
})();

const outDir = "C:/Users/LENOVO/projects/ai-education-pptx";
pres.writeFile({ fileName: outDir + "/AI-Education-Hermes.pptx" }).then(() => {
  console.log("DONE: " + outDir + "/AI-Education-Hermes.pptx");
  console.log("Slides: 6 | 16:9");
}).catch(e => { console.error(e); process.exit(1); });
