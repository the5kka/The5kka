const fs = require('fs');
const path = require('path');

const root = __dirname;
const sourcePath = path.join(root, 'gallery_60.html');
const outputPath = path.join(root, 'OJT_디자인_60종_핸드폰용.html');
let html = fs.readFileSync(sourcePath, 'utf8');

const embeddedScreens = [null];
for (let i = 1; i <= 60; i += 1) {
  const name = `design_${String(i).padStart(2, '0')}.png`;
  const data = fs.readFileSync(path.join(root, 'screens', name)).toString('base64');
  embeddedScreens.push(`data:image/png;base64,${data}`);
}

const lucide = fs.readFileSync(path.join(root, 'lucide.min.js'), 'utf8');
const mobileCss = `
  @media (max-width: 760px) {
    body { min-width: 0; }
    .topbar { position: static; padding: 14px; flex-direction: column; align-items: stretch; gap: 12px; }
    .title { min-width: 0; }
    .title h1 { font-size: 18px; }
    .tabs { width: 100%; margin: 0; display: grid; grid-template-columns: 1fr; }
    .tab { width: 100%; justify-content: center; }
    .summary { padding: 12px; grid-template-columns: 1fr 1fr; gap: 7px; }
    .stat { min-height: 50px; padding: 8px; }
    .stat b { font-size: 20px; }
    main { padding: 12px 10px 28px; }
    .section-head { align-items: start; flex-direction: column; gap: 3px; }
    .screen-grid { grid-template-columns: 1fr; gap: 12px; }
    .palette-grid, .button-grid { grid-template-columns: 1fr 1fr; gap: 8px; }
    .palette-card, .button-card { min-height: 122px; padding: 8px; }
    .swatches, .button-stage { height: 68px; }
    .sample-button { min-width: 112px; padding: 0 8px; font-size: 9px; }
    dialog { width: 100vw; max-width: none; max-height: 100vh; border-radius: 0; }
    .viewer-body { padding: 6px; grid-template-columns: 30px 1fr 30px; gap: 4px; }
    .nav { width: 30px; height: 58px; }
    .viewer-body img { max-height: calc(100vh - 70px); }
  }
`;

html = html.replace('</style>', `${mobileCss}</style>`);
html = html.replace('<title>OJT Exam Maker 디자인 60종</title>', '<title>OJT Exam Maker 모바일 디자인 60종</title>');
html = html.replace('const families = [', `const embeddedScreens = ${JSON.stringify(embeddedScreens)};\n    const families = [`);
html = html.replace(
  "document.getElementById('viewerImage').src=`screens/design_${n}.png`;",
  "document.getElementById('viewerImage').src=embeddedScreens[current];"
);
html = html.split('screens/design_${n}.png').join('${embeddedScreens[index]}');
html = html.replace('<script src="lucide.min.js"></script>', `<script>${lucide}</script>`);
fs.writeFileSync(outputPath, html, 'utf8');
console.log(outputPath);
