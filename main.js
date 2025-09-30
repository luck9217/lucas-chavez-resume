// Minimal helpers
const q = (sel, root = document) => root.querySelector(sel);
const qa = (sel, root = document) => Array.from(root.querySelectorAll(sel));

// Build a minimal styled HTML for Word export
function buildWordHTML(title, bodyHTML) {
  const styles = `body{font-family:Arial,Helvetica,sans-serif;line-height:1.45;color:#111;padding:16px}
    h1{font-size:22px;margin:0 0 8px}
    h2{font-size:16px;margin:18px 0 6px;text-transform:uppercase;color:#555}
    h3{font-size:14px;margin:12px 0 4px}
    p,li{font-size:14px;margin:6px 0}
    ul{padding-left:20px}
    .meta{color:#555;font-size:13px;margin-bottom:4px}
    .hr{height:1px;background:#e5e7eb;margin:16px 0}
    .resume-header{margin-bottom:8px}
    .resume-header p{margin:2px 0}`;
  return `<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>${title}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <style>${styles}</style></head><body>${bodyHTML}</body></html>`;
}

// Export resume section with a fixed contact header to Word
function exportToWord(
  sectionId = "resume",
  filename = "Lucas_Chavez_Resume.doc"
) {
  const section = document.getElementById(sectionId);
  if (!section) return;
  const content = q(".content", section) || section;

  const container = document.createElement("div");
  // Fixed resume header required at the top of the Word file
  const resumeHeader = document.createElement("div");
  resumeHeader.className = "resume-header";
  resumeHeader.innerHTML = `
    <h1>Lucas Chavez</h1>
    <p>LUCK9217@gmail.com</p>
    <p>+61 467520280</p>
    <p>LinkedIn: https://www.linkedin.com/in/lucas-chavez/</p>
    <div class="hr"></div>
  `;

  // Append header then cloned resume content
  container.appendChild(resumeHeader);
  container.appendChild(content.cloneNode(true));

  // Remove any toolbars in the clone
  qa(".toolbar", container).forEach((el) => el.remove());

  const html = buildWordHTML(filename, container.innerHTML);
  const blob = new Blob(["\ufeff", html], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
