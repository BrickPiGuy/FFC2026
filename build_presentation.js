const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
const SW = 13.333333;
const SH = 7.5;
const FONT = 'Aptos';
const MONO = 'Aptos Mono';
const OUT = 'ECPI_Augmented_Craftsman_Faculty_Conference_2026.pptx';

const C = {
  bg: '081A2B',
  panel: '102A40',
  panel2: '15364F',
  line: '294A60',
  white: 'F5F8FA',
  muted: 'A8BDCD',
  teal: '49D2BF',
  blue: '7DAFFF',
  amber: 'F4BD62',
  coral: 'EC8C80',
};

pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'ECPI University';
pptx.subject = 'November 2026 Faculty Conference';
pptx.title = 'The Future of Software Development at ECPI';
pptx.company = 'ECPI University';
pptx.lang = 'en-US';
pptx.theme = {
  headFontFace: FONT,
  bodyFontFace: FONT,
  lang: 'en-US',
};
pptx.defineSlideMaster({
  title: 'BASE',
  background: { color: C.bg },
  objects: [
    { rect: { x: 0, y: 0, w: SW, h: 0.055, line: { color: C.teal, transparency: 100 }, fill: { color: C.teal } } },
  ],
  slideNumber: { x: 12.18, y: 7.06, w: 0.5, h: 0.18, color: C.muted, fontFace: FONT, fontSize: 8, align: 'right', margin: 0 },
});

function addText(slide, text, x, y, w, h, options = {}) {
  slide.addText(text, {
    x, y, w, h,
    margin: 0,
    fontFace: FONT,
    fontSize: 14,
    color: C.white,
    breakLine: false,
    valign: 'mid',
    ...options,
  });
}

function rect(slide, x, y, w, h, fill, options = {}) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    line: { color: fill, transparency: 100 },
    fill: { color: fill },
    ...options,
  });
}

function roundRect(slide, x, y, w, h, fill, options = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.08,
    line: { color: fill, transparency: 100 },
    fill: { color: fill },
    ...options,
  });
}

function line(slide, x1, y1, x2, y2, color = C.line, width = 1, transparency = 0) {
  slide.addShape(pptx.ShapeType.line, {
    x: x1, y: y1, w: x2 - x1, h: y2 - y1,
    line: { color, width, transparency, beginArrowType: 'none', endArrowType: 'none' },
  });
}

function dot(slide, x, y, diameter, color) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x, y, w: diameter, h: diameter,
    line: { color, transparency: 100 },
    fill: { color },
  });
}

function addHeader(slide, section, index, title, subtitle) {
  addText(slide, `ECPI UNIVERSITY   /   FACULTY CONFERENCE   /   ${section.toUpperCase()}`,
    0.72, 0.31, 10.4, 0.22, { fontSize: 9, bold: true, charSpacing: 1.5, color: C.teal });
  addText(slide, `${String(index).padStart(2, '0')}  /  08`, 11.85, 0.31, 0.78, 0.22,
    { fontSize: 9, bold: true, align: 'right', color: C.muted });
  addText(slide, title, 0.72, 0.76, 11.9, 0.61,
    { fontSize: 29, bold: true, breakLine: false });
  if (subtitle) {
    addText(slide, subtitle, 0.74, 1.43, 11.3, 0.35,
      { fontSize: 13, color: C.muted });
  }
}

function addNotes(slide, message) {
  slide.addNotes(message.trim().split('\n').map((part) => part.trim()).filter(Boolean));
}

function addGlyph(slide, kind, x, y, color) {
  const thin = { color, width: 1.5 };
  if (kind === 'ai') {
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.04, y: y + 0.05, w: 0.45, h: 0.34, rectRadius: 0.04, line: thin, fill: { color: C.panel, transparency: 100 } });
    line(slide, x + 0.14, y + 0.17, x + 0.22, y + 0.23, color, 1.5);
    line(slide, x + 0.22, y + 0.23, x + 0.14, y + 0.29, color, 1.5);
    line(slide, x + 0.29, y + 0.29, x + 0.39, y + 0.17, color, 1.5);
  } else if (kind === 'git') {
    line(slide, x + 0.1, y + 0.12, x + 0.1, y + 0.38, color, 1.5);
    line(slide, x + 0.1, y + 0.24, x + 0.39, y + 0.12, color, 1.5);
    line(slide, x + 0.39, y + 0.12, x + 0.39, y + 0.37, color, 1.5);
    dot(slide, x + 0.035, y + 0.08, 0.13, color);
    dot(slide, x + 0.325, y + 0.055, 0.13, color);
    dot(slide, x + 0.325, y + 0.33, 0.13, color);
    dot(slide, x + 0.035, y + 0.31, 0.13, color);
  } else if (kind === 'pipeline') {
    [0, 1, 2].forEach((index) => rect(slide, x + index * 0.17, y + 0.15, 0.12, 0.19, color));
    line(slide, x + 0.09, y + 0.39, x + 0.44, y + 0.39, color, 1.3);
  } else if (kind === 'project') {
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.05, y: y + 0.04, w: 0.36, h: 0.35, rectRadius: 0.03, line: thin, fill: { color: C.panel, transparency: 100 } });
    line(slide, x + 0.13, y + 0.15, x + 0.34, y + 0.15, color, 1.4);
    line(slide, x + 0.13, y + 0.23, x + 0.3, y + 0.23, color, 1.4);
    line(slide, x + 0.13, y + 0.31, x + 0.25, y + 0.31, color, 1.4);
  } else if (kind === 'credential') {
    slide.addShape(pptx.ShapeType.ellipse, { x: x + 0.11, y: y + 0.035, w: 0.32, h: 0.32, line: thin, fill: { color: C.panel, transparency: 100 } });
    dot(slide, x + 0.22, y + 0.145, 0.1, color);
    line(slide, x + 0.17, y + 0.34, x + 0.15, y + 0.44, color, 1.4);
    line(slide, x + 0.37, y + 0.34, x + 0.39, y + 0.44, color, 1.4);
  } else {
    [0, 1, 2, 3].forEach((index) => {
      slide.addShape(pptx.ShapeType.rect, {
        x: x + (index % 2) * 0.21 + 0.05,
        y: y + Math.floor(index / 2) * 0.21 + 0.05,
        w: 0.14, h: 0.14,
        line: { color, width: 1.2 },
        fill: { color: index === 0 ? color : C.panel, transparency: index === 0 ? 0 : 100 },
      });
    });
  }
}

const slides = [];

// Slide 1
{
  const slide = pptx.addSlide('BASE');
  slides.push(slide);
  addText(slide, 'ECPI UNIVERSITY   /   NOVEMBER 2026 FACULTY CONFERENCE',
    0.78, 0.56, 6.85, 0.24, { fontSize: 9, bold: true, charSpacing: 1.4, color: C.teal });
  addText(slide, 'The Future of\nSoftware Development\nat ECPI',
    0.78, 1.35, 7.03, 2.52, { fontSize: 35, bold: true, breakLine: false, valign: 'top', lineSpacingMultiple: 0.94 });
  rect(slide, 0.8, 4.17, 0.62, 0.055, C.teal);
  addText(slide, 'The Augmented Craftsman Model',
    0.8, 4.42, 6.7, 0.44, { fontSize: 21, bold: true, color: C.teal });
  addText(slide, 'Preserving the craft. Embracing the tools.\nPreparing the next generation.',
    0.8, 5.13, 6.85, 0.88, { fontSize: 16, color: C.muted, valign: 'top', breakLine: false });
  addText(slide, 'NOVEMBER 2026   /   FACULTY CONFERENCE',
    0.8, 6.57, 6.4, 0.25, { fontSize: 9, bold: true, charSpacing: 1.1, color: C.muted });

  roundRect(slide, 8.05, 0.88, 4.55, 5.62, C.panel);
  [0, 1, 2].forEach((i) => dot(slide, 8.39 + i * 0.2, 1.16, 0.085, [C.teal, C.blue, C.amber][i]));
  addText(slide, 'A PRACTICE, NOT A SHORTCUT', 9.05, 1.1, 3.08, 0.19,
    { fontSize: 8, bold: true, charSpacing: 1.1, color: C.muted, align: 'right' });
  line(slide, 8.4, 1.57, 12.25, 1.57, C.line, 1);
  addText(slide, 'intent  /  design  /  build', 8.43, 1.94, 3.7, 0.34,
    { fontFace: MONO, fontSize: 12, color: C.muted });
  addText(slide, 'judgment()', 8.43, 2.58, 2.15, 0.43,
    { fontFace: MONO, fontSize: 20, bold: true, color: C.white });
  rect(slide, 10.59, 2.75, 0.68, 0.045, C.teal);
  addText(slide, 'test()', 9.0, 3.36, 1.5, 0.39,
    { fontFace: MONO, fontSize: 18, bold: true, color: C.blue });
  rect(slide, 10.5, 3.53, 0.5, 0.04, C.blue);
  addText(slide, 'refine()', 9.57, 4.13, 1.7, 0.39,
    { fontFace: MONO, fontSize: 18, bold: true, color: C.amber });
  rect(slide, 11.3, 4.3, 0.55, 0.04, C.amber);
  line(slide, 8.62, 3.1, 8.62, 4.38, C.line, 1.5);
  [3.02, 3.66, 4.32].forEach((y, i) => dot(slide, 8.55, y, 0.14, [C.teal, C.blue, C.amber][i]));
  line(slide, 8.4, 4.93, 12.25, 4.93, C.line, 1);
  addText(slide, 'CAPABILITY GROWS THROUGH PRACTICE', 8.43, 5.28, 3.85, 0.24,
    { fontSize: 9, bold: true, charSpacing: 0.8, color: C.teal });
  addText(slide, 'Human craft, strengthened by the right tools.',
    8.43, 5.64, 3.75, 0.52, { fontSize: 12, color: C.muted, valign: 'top' });

  addNotes(slide, `Key message: ECPI is shaping a software development program that pairs enduring engineering craft with thoughtful use of modern tools.\n\nSuggested talking points: Welcome faculty and colleagues. The phrase “Augmented Craftsman” captures an educational direction, not a promise that a tool can do the work of learning. We want students to understand what they build, make sound design choices, test their assumptions, and use AI where it helps them work with greater capability. The visual on the right is a working cycle: intent, judgment, testing, refinement. Each step still depends on a developer making decisions. This conference is a chance to share the vision, show real development activity, and invite the faculty expertise that will make the program durable.\n\nTransition: First, let’s connect that vision to what is changing in the practice of software development.`);
}

// Slide 2
{
  const slide = pptx.addSlide('BASE');
  slides.push(slide);
  addHeader(slide, 'The case for change', 2, 'Why We Are Evolving', 'Software development education must evolve with the profession.');
  const cards = [
    { x: 0.73, n: '01', color: C.teal, title: 'AI changes\nthe workflow', copy: 'Developers are learning to work with AI-assisted tools, not just write code by hand.' },
    { x: 4.81, n: '02', color: C.blue, title: 'Judgment matters\nmore than output', copy: 'Students need to reason about solutions, evaluate trade-offs, and solve real problems.' },
    { x: 8.89, n: '03', color: C.amber, title: 'Delivery is\na team practice', copy: 'Modern work brings together collaboration, version control, testing, and automation.' },
  ];
  for (const card of cards) {
    roundRect(slide, card.x, 2.03, 3.72, 2.45, C.panel);
    addText(slide, card.n, card.x + 0.26, 2.27, 0.7, 0.5,
      { fontSize: 25, bold: true, color: card.color });
    rect(slide, card.x + 0.27, 2.92, 0.49, 0.045, card.color);
    addText(slide, card.title, card.x + 0.27, 3.12, 3.1, 0.66,
      { fontSize: 19, bold: true, valign: 'top' });
    addText(slide, card.copy, card.x + 0.27, 3.86, 3.14, 0.44,
      { fontSize: 10.5, color: C.muted, valign: 'top' });
  }
  roundRect(slide, 0.73, 4.93, 11.88, 1.16, C.panel2);
  rect(slide, 0.99, 5.19, 0.055, 0.62, C.teal);
  addText(slide, 'The goal is not simply to teach students new tools.',
    1.24, 5.11, 10.8, 0.32, { fontSize: 17, color: C.muted });
  addText(slide, 'It is to develop stronger software developers.',
    1.24, 5.5, 10.8, 0.38, { fontSize: 20, bold: true, color: C.white });
  addNotes(slide, `Key message: The program is evolving in response to changes in practice, while keeping student capability at the center.\n\nSuggested talking points: Three forces make this a timely conversation. AI-assisted development changes how people approach a task, but it does not remove the need to know whether a result is correct, useful, or appropriate. Employers need graduates who can frame a problem, explain a choice, and assess a proposed solution; code production alone is not the whole job. And software reaches users through coordinated work: a developer collaborates, manages changes, tests, and improves the product. Those practices are learned by doing, with guidance and feedback. We are not proposing a curriculum organized around the novelty of a tool. Our aim is to help students build stronger and more flexible software development habits.\n\nTransition: That educational goal is the heart of the Augmented Craftsman model.`);
}

// Slide 3
{
  const slide = pptx.addSlide('BASE');
  slides.push(slide);
  addHeader(slide, 'Educational philosophy', 3, 'The Augmented Craftsman Model', 'AI can extend capability; the developer remains responsible for the work.');
  const nodes = [
    { x: 0.65, y: 2.02, title: 'Engineering fundamentals', sub: 'Understand the systems we build.', color: C.teal },
    { x: 0.42, y: 3.58, title: 'Critical thinking', sub: 'Frame problems and weigh trade-offs.', color: C.blue },
    { x: 0.65, y: 5.12, title: 'Practical craftsmanship', sub: 'Write clear, maintainable software.', color: C.amber },
    { x: 10.02, y: 2.02, title: 'Responsible AI use', sub: 'Use assistance with discernment.', color: C.teal },
    { x: 10.25, y: 3.58, title: 'Collaborative practice', sub: 'Build through shared workflows.', color: C.blue },
    { x: 10.02, y: 5.12, title: 'Continuous learning', sub: 'Adapt as tools and needs change.', color: C.amber },
  ];
  const centerX = 6.665;
  const centerY = 3.72;
  for (const node of nodes) {
    line(slide, centerX, centerY, node.x + 1.33, node.y + 0.48, C.line, 1.1, 8);
  }
  for (const node of nodes) {
    roundRect(slide, node.x, node.y, 2.66, 0.96, C.panel);
    rect(slide, node.x + 0.18, node.y + 0.19, 0.045, 0.56, node.color);
    addText(slide, node.title, node.x + 0.36, node.y + 0.18, 2.12, 0.27,
      { fontSize: 12.5, bold: true });
    addText(slide, node.sub, node.x + 0.36, node.y + 0.5, 2.11, 0.24,
      { fontSize: 9, color: C.muted });
  }
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 5.38, y: 2.44, w: 2.57, h: 2.57,
    line: { color: C.teal, width: 1.2, transparency: 10 },
    fill: { color: C.bg },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 5.62, y: 2.68, w: 2.09, h: 2.09,
    line: { color: C.line, width: 1 },
    fill: { color: C.panel },
  });
  addText(slide, 'AUGMENTED', 5.85, 3.26, 1.63, 0.22,
    { fontSize: 8, bold: true, charSpacing: 1.3, align: 'center', color: C.teal });
  addText(slide, 'Craftsman', 5.63, 3.58, 2.08, 0.41,
    { fontSize: 19, bold: true, align: 'center' });
  addText(slide, 'Judgment leads.\nTools extend.', 5.8, 4.09, 1.73, 0.44,
    { fontSize: 10, align: 'center', color: C.muted, breakLine: false });
  addNotes(slide, `Key message: The model treats technical knowledge, thoughtful practice, and responsible augmentation as one integrated kind of professional development.\n\nSuggested talking points: It may help to think of these capabilities as mutually reinforcing, rather than as a checklist of separate subjects. Strong fundamentals help a student evaluate code, including code proposed by AI. Critical thinking gives them a way to ask whether a solution addresses the actual problem. Craftsmanship turns a promising idea into software another person can understand and maintain. Collaborative practices let those skills work across a team. Responsible AI use requires students to be transparent, attentive to quality, and accountable for what they submit. And because the field will continue to change, adaptability is itself part of the craft. The center of this model is a person exercising judgment; AI is a tool, not a replacement for that judgment.\n\nTransition: Those principles become more tangible in the learning experiences we are proposing.`);
}

// Slide 4
{
  const slide = pptx.addSlide('BASE');
  slides.push(slide);
  addHeader(slide, 'Learning experience', 4, 'What Makes This Program Different?', 'A practical pathway connects tools, disciplined practice, and student judgment.');
  const features = [
    { kind: 'ai', title: 'Applied AI', copy: 'Practice using AI throughout the development process.', color: C.teal },
    { kind: 'git', title: 'Collaborative development', copy: 'Build fluency with Git and shared workflows.', color: C.blue },
    { kind: 'pipeline', title: 'CI/CD + DevOps', copy: 'Connect automation, testing, and delivery.', color: C.amber },
    { kind: 'project', title: 'Hands-on projects', copy: 'Learn by creating, reviewing, and refining.', color: C.blue },
    { kind: 'credential', title: 'Certification pathway', copy: 'Integrate industry certifications into the pathway.', color: C.teal },
    { kind: 'craft', title: 'Pragmatic engineering', copy: 'Pair creativity with sound technical judgment.', color: C.amber },
  ];
  const xs = [0.72, 4.81, 8.9];
  const ys = [1.95, 4.2];
  for (let index = 0; index < features.length; index += 1) {
    const feature = features[index];
    const x = xs[index % 3];
    const y = ys[Math.floor(index / 3)];
    roundRect(slide, x, y, 3.72, 1.93, C.panel);
    addGlyph(slide, feature.kind, x + 0.24, y + 0.22, feature.color);
    addText(slide, String(index + 1).padStart(2, '0'), x + 3.08, y + 0.3, 0.34, 0.19,
      { fontSize: 8.5, bold: true, align: 'right', color: C.muted });
    addText(slide, feature.title, x + 0.25, y + 0.83, 3.18, 0.37,
      { fontSize: 16, bold: true });
    addText(slide, feature.copy, x + 0.25, y + 1.29, 3.13, 0.36,
      { fontSize: 10.5, color: C.muted, valign: 'top' });
  }
  addText(slide, 'A proposed learning direction; specific curriculum and certification requirements remain subject to review.',
    0.76, 6.53, 11.7, 0.22, { fontSize: 9, color: C.muted, italic: true });
  addNotes(slide, `Key message: The program is designed to connect modern development practices with applied, project-based learning and an intentional professional pathway.\n\nSuggested talking points: These six features describe the shape of the proposed student experience. Students should encounter AI in the context of real software work, not as a disconnected topic. Version control and collaboration make that work visible and teachable. Automation, testing, and continuous delivery help students see how individual code changes become reliable releases. Projects let them put those practices together and learn from feedback. Industry certifications are intended to be integrated into the educational pathway, though we should be careful not to imply that particular certification requirements have been finalized here. Across the experience, pragmatic engineering and creativity belong together: students should be able to explore possibilities while explaining and defending their technical choices.\n\nTransition: With that proposed experience in view, we can look at the current development work, beginning with Phase One.`);
}

// Slide 5
{
  const slide = pptx.addSlide();
  slides.push(slide);
  slide.addImage({ path: 'Phase_One_Progress_Infographic.png', x: 0, y: 0, w: SW, h: SH });
  addNotes(slide, `Key message: Phase One shows meaningful development progress, while distinguishing published courses from courses still in development or quality control.\n\nSuggested talking points: The approved infographic is the visual anchor for this update. The presenter has manually verified these figures against the Phase One workbook. It depicts 30 courses in Phase One scope and 14 as published. That published label matters: publication is a deliverable, whereas the courses shown at 90 percent are still at the quality-control handoff stage, not published courses. The graphic also shows a development pipeline with a range of work underway and one course on hold for subject-matter-expert availability. The selected spotlight courses are illustrative rather than a full roster.\n\nTransition: Phase Two is the next wave, and its visible pipeline tells a different story.`);
}

// Slide 6
{
  const slide = pptx.addSlide();
  slides.push(slide);
  slide.addImage({ path: 'Phase_Two_Progress_Infographic.png', x: 0, y: 0, w: SW, h: SH });
  addNotes(slide, `Key message: Phase Two is a growing development pipeline, not a set of published courses; the selected spotlight makes that distinction clear.\n\nSuggested talking points: The presenter has manually verified these figures against the Phase Two workbook. The graphic depicts 20 courses in scope: four in development, thirteen awaiting resources, three not started, and none published. Those categories describe different kinds of work and should not be collapsed into a single completion figure. The spotlight shows SDC130, Database Systems and SQL, at 90 percent, and SDC305, Software Architecture Using AI Tools, at 85 percent. Both are shown as in development, not as published deliverables. SDC130 also appears in the Phase One graphic, so it must not be counted twice if we later summarize the program portfolio. The resource queue is part of the honest picture: a visible pipeline provides a basis for planning, not a claim that the queued work is complete.\n\nTransition: These two phases point toward a sequence of deliberate next steps, rather than a set of invented deadline dates.`);
}

// Slide 7
{
  const slide = pptx.addSlide('BASE');
  slides.push(slide);
  addHeader(slide, 'Next steps', 7, 'Our Development Roadmap', 'A sequence of milestones, with timing set as the work and resources are confirmed.');
  const steps = [
    { x: 1.27, title: 'Develop + refine', copy: 'Create course content\nand learning activities.', color: C.teal },
    { x: 3.97, title: 'Review + quality', copy: 'Strengthen accuracy,\nclarity, and instruction.', color: C.blue },
    { x: 6.67, title: 'Prepare faculty', copy: 'Align on course goals\nand teaching practice.', color: C.amber },
    { x: 9.37, title: 'Release + implement', copy: 'Bring reviewed courses\ninto the pathway.', color: C.teal },
    { x: 12.07, title: 'Listen + improve', copy: 'Use feedback to guide\ncontinuous improvement.', color: C.blue },
  ];
  line(slide, 1.52, 3.08, 11.82, 3.08, C.line, 2);
  for (let index = 0; index < steps.length; index += 1) {
    const step = steps[index];
    dot(slide, step.x - 0.14, 2.94, 0.28, step.color);
    addText(slide, `0${index + 1}`, step.x - 0.39, 2.24, 0.78, 0.23,
      { fontSize: 9, bold: true, align: 'center', charSpacing: 0.8, color: step.color });
    addText(slide, step.title, step.x - 1.07, 3.42, 2.14, 0.49,
      { fontSize: 13, bold: true, align: 'center', valign: 'top' });
    addText(slide, step.copy, step.x - 1.15, 4.04, 2.3, 0.54,
      { fontSize: 10, color: C.muted, align: 'center', valign: 'top', breakLine: false });
  }
  roundRect(slide, 0.78, 5.34, 11.77, 0.96, C.panel2);
  addText(slide, 'A deliberate, phased approach to building a sustainable modern software development program.',
    1.08, 5.59, 11.16, 0.39, { fontSize: 17, bold: true, align: 'center' });
  addText(slide, 'Sequence shown; no milestone dates are asserted.',
    0.83, 6.55, 11.5, 0.2, { fontSize: 9, color: C.muted, italic: true, align: 'center' });
  addNotes(slide, `Key message: A sustainable program depends on a sequence of development, quality review, faculty preparation, implementation, and learning from feedback.\n\nSuggested talking points: The roadmap is intentionally organized by sequence rather than calendar date. We have not been given documented milestone dates, and a useful conference update should not suggest that tentative timing is approved. First, content is developed and refined. Then instructional and quality reviews give us a chance to catch issues while the material is still being improved. Faculty preparation helps connect course goals to the teaching experience. Course release and implementation follow those steps rather than substitute for them. Finally, feedback from instructors and students should inform continuous improvement. The sequence also gives faculty several meaningful points of participation. The point is not to promise a fixed schedule today; it is to show a disciplined approach for bringing the work forward responsibly.\n\nTransition: That process will be strongest when faculty help shape both the learning experience and the feedback loop.`);
}

// Slide 8
{
  const slide = pptx.addSlide('BASE');
  slides.push(slide);
  addHeader(slide, 'Faculty partnership', 8, 'Building the Future Together', 'The transformation is about the people our students are becoming.');
  addText(slide, 'Equip students with the judgment, confidence, craftsmanship, and adaptability to succeed in an evolving profession.',
    0.78, 1.96, 11.63, 0.72,
    { fontSize: 17, color: C.muted, valign: 'top', breakLine: false });

  const contributions = [
    { x: 0.78, color: C.teal, title: 'Share expertise', copy: 'Bring instructional insight\nand practical experience.' },
    { x: 4.82, color: C.blue, title: 'Shape the review', copy: 'Contribute to course review\nand ongoing improvement.' },
    { x: 8.86, color: C.amber, title: 'Coach judgment', copy: 'Help students pair new skills\nwith professional discernment.' },
  ];
  for (const item of contributions) {
    rect(slide, item.x, 3.0, 0.48, 0.045, item.color);
    addText(slide, item.title, item.x, 3.25, 3.6, 0.35,
      { fontSize: 15, bold: true });
    addText(slide, item.copy, item.x, 3.71, 3.55, 0.53,
      { fontSize: 11, color: C.muted, valign: 'top', breakLine: false });
  }
  roundRect(slide, 0.78, 4.72, 11.79, 1.42, C.panel2);
  addText(slide, '“The future of software development is not AI versus developers.',
    1.12, 4.98, 11.15, 0.36, { fontSize: 17, bold: true, align: 'center' });
  addText(slide, 'It is better developers, augmented by AI.”',
    1.12, 5.43, 11.15, 0.37, { fontSize: 20, bold: true, align: 'center', color: C.teal });
  addText(slide, 'DISCUSSION', 0.83, 6.6, 2.2, 0.2,
    { fontSize: 9, bold: true, charSpacing: 1.3, color: C.teal });
  addNotes(slide, `Key message: Faculty partnership is essential to developing students who can combine sound judgment with emerging technical capability.\n\nSuggested talking points: This change is about more than new courses or new technologies. It is about equipping students with the confidence to engage with unfamiliar problems, the craft to build well, and the adaptability to continue learning. Faculty expertise is central to that effort. Your practical and instructional insight can make course materials more relevant and more teachable. Participating in review helps ensure that the work is clear, rigorous, and useful to students. And in the classroom, faculty can help students develop professional judgment alongside technical skill: asking what a tool did, whether the result is sound, and what responsibility the developer has for the outcome. The invitation is to help shape this together. The central idea is not AI versus developers; it is better developers, augmented by AI. I would welcome your questions and discussion.\n\nTransition: Open the floor for discussion.`);
}

if (slides.length !== 8) {
  throw new Error(`Expected exactly 8 slides, generated ${slides.length}.`);
}

pptx.writeFile({ fileName: OUT });