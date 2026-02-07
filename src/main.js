/* global katex, markdownit, TurndownService */

const STORAGE_KEY = "typora-clone-wysiwyg-html-v1";
const SAVE_DEBOUNCE_MS = 120;
const MATH_PASS_DEBOUNCE_MS = 70;

const DEFAULT_HTML = `
<h1>Typora Clone WYSIWYG</h1>
<p>Type naturally in rich text mode. Math uses <code>$...$</code> and <code>$$...$$</code>.</p>
<p>Inline: $E = mc^2$ and $e^{i\\pi}+1=0$</p>
<p>Block:</p>
<p>$$\\int_{-\\infty}^{\\infty} e^{-x^2} \\, dx = \\sqrt{\\pi}$$</p>
`;

const editor = document.querySelector("#editor");
const latencyEl = document.querySelector("#latency");
const toolbar = document.querySelector(".toolbar");
const importInput = document.querySelector("#import-md");
const mathPreview = document.querySelector("#math-preview");
const mathPreviewContent = document.querySelector("#math-preview-content");

if (!editor || !latencyEl || !toolbar || !importInput || !mathPreview || !mathPreviewContent) {
  throw new Error("Missing required DOM elements");
}

const md = markdownit({ html: true, linkify: true, breaks: false });
const turndown = new TurndownService({ headingStyle: "atx", codeBlockStyle: "fenced" });

turndown.addRule("mathToken", {
  filter: (node) => node.nodeType === Node.ELEMENT_NODE && node.classList.contains("math-token"),
  replacement: (_, node) => {
    const tex = node.getAttribute("data-tex") || "";
    const display = node.getAttribute("data-display") === "true";
    return display ? `\n\n$$\n${tex}\n$$\n\n` : `$${tex}$`;
  }
});

let mathPassTimer = null;
let saveTimer = null;
let activeMathEdit = null;

editor.innerHTML = localStorage.getItem(STORAGE_KEY) || DEFAULT_HTML;
runMathPass(true);
updateMathPreview();

function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    localStorage.setItem(STORAGE_KEY, editor.innerHTML);
  }, SAVE_DEBOUNCE_MS);
}

function scheduleMathPass(aggressive = false) {
  clearTimeout(mathPassTimer);
  mathPassTimer = setTimeout(() => runMathPass(aggressive), MATH_PASS_DEBOUNCE_MS);
}

function createMathToken(tex, displayMode) {
  const span = document.createElement("span");
  span.className = "math-token";
  span.setAttribute("contenteditable", "false");
  span.setAttribute("data-tex", tex);
  span.setAttribute("data-display", displayMode ? "true" : "false");

  try {
    span.innerHTML = katex.renderToString(tex, {
      throwOnError: false,
      strict: "ignore",
      displayMode
    });
  } catch {
    span.textContent = displayMode ? `$$${tex}$$` : `$${tex}$`;
  }

  return span;
}

function currentSelectionTextNode() {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) {
    return null;
  }

  const anchor = sel.anchorNode;
  if (!anchor) {
    return null;
  }

  return anchor.nodeType === Node.TEXT_NODE ? anchor : anchor.firstChild;
}

function runMathPass(aggressive) {
  const startedAt = performance.now();
  const activeTextNode = aggressive ? null : currentSelectionTextNode();
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  const toReplace = [];

  while (walker.nextNode()) {
    const node = walker.currentNode;
    const parent = node.parentElement;
    if (!parent) {
      continue;
    }

    if (parent.closest(".math-token, code, pre, script, style")) {
      continue;
    }

    if (!aggressive && activeTextNode && node === activeTextNode) {
      continue;
    }

    const text = node.nodeValue || "";
    if (!text.includes("$")) {
      continue;
    }

    if (!/(\$\$[\s\S]+?\$\$)|(\$[^$\n]+?\$)/.test(text)) {
      continue;
    }

    toReplace.push(node);
  }

  for (const node of toReplace) {
    const value = node.nodeValue || "";
    const fragment = document.createDocumentFragment();
    const regex = /(\$\$([\s\S]+?)\$\$)|(\$([^$\n]+?)\$)/g;
    let last = 0;
    let match = regex.exec(value);

    while (match) {
      if (match.index > last) {
        fragment.appendChild(document.createTextNode(value.slice(last, match.index)));
      }

      const isBlock = Boolean(match[1]);
      const tex = (isBlock ? match[2] : match[4] || "").trim();
      if (tex.length > 0) {
        fragment.appendChild(createMathToken(tex, isBlock));
      } else {
        fragment.appendChild(document.createTextNode(match[0]));
      }

      last = regex.lastIndex;
      match = regex.exec(value);
    }

    if (last < value.length) {
      fragment.appendChild(document.createTextNode(value.slice(last)));
    }

    node.replaceWith(fragment);
  }

  latencyEl.textContent = `Math pass: ${(performance.now() - startedAt).toFixed(1)} ms`;
  if (activeMathEdit && !editor.contains(activeMathEdit.textNode)) {
    activeMathEdit = null;
  }
  scheduleSave();
  updateMathPreview();
}

function insertMathByPrompt() {
  const tex = window.prompt("LaTeX expression (without delimiters):", "x^2 + y^2 = z^2");
  if (tex == null) {
    return;
  }

  const isBlock = window.confirm("Render as block math?\nOK = block ($$...$$), Cancel = inline ($...$)");
  const raw = isBlock ? `$$${tex}$$` : `$${tex}$`;

  document.execCommand("insertText", false, raw);
  scheduleMathPass(true);
}

function applyHeading(level) {
  document.execCommand("formatBlock", false, level === 1 ? "h1" : "h2");
}

function sanitizeImportedHtml(html) {
  const wrapper = document.createElement("div");
  wrapper.innerHTML = html;
  wrapper.querySelectorAll("script, iframe, object, embed").forEach((el) => el.remove());
  return wrapper.innerHTML;
}

function exportMarkdown() {
  const cloned = editor.cloneNode(true);
  cloned.querySelectorAll(".math-token").forEach((token) => {
    const tex = token.getAttribute("data-tex") || "";
    const display = token.getAttribute("data-display") === "true";
    token.replaceWith(document.createTextNode(display ? `$$${tex}$$` : `$${tex}$`));
  });

  const markdown = turndown.turndown(cloned.innerHTML).trimEnd() + "\n";
  const blob = new Blob([markdown], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = "document.md";
  anchor.click();
  URL.revokeObjectURL(url);
}

function isMathTokenNode(node) {
  return node instanceof HTMLElement && node.classList.contains("math-token");
}

function mathRawForToken(token) {
  const tex = token.getAttribute("data-tex") || "";
  const display = token.getAttribute("data-display") === "true";
  const delimiterSize = display ? 2 : 1;
  return {
    tex,
    display,
    delimiterSize,
    raw: display ? `$$${tex}$$` : `$${tex}$`
  };
}

function setCaret(node, offset) {
  const range = document.createRange();
  range.setStart(node, Math.max(0, Math.min(offset, node.nodeValue ? node.nodeValue.length : 0)));
  range.collapse(true);

  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  selection.removeAllRanges();
  selection.addRange(range);
}

function setCaretAfterNode(node) {
  const range = document.createRange();
  range.setStartAfter(node);
  range.collapse(true);

  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  selection.removeAllRanges();
  selection.addRange(range);
}

function setCaretBeforeNode(node) {
  const range = document.createRange();
  range.setStartBefore(node);
  range.collapse(true);

  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  selection.removeAllRanges();
  selection.addRange(range);
}

function setCaretBeforeReferenceNode(parent, node) {
  const range = document.createRange();
  range.setStartBefore(node);
  range.collapse(true);

  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  selection.removeAllRanges();
  selection.addRange(range);
  if (parent instanceof HTMLElement) {
    parent.focus();
  }
}

function getCollapsedEditorRange() {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0 || !selection.isCollapsed) {
    return null;
  }

  const range = selection.getRangeAt(0);
  if (!editor.contains(range.startContainer)) {
    return null;
  }

  return range;
}

function adjacentSignificantSibling(node, direction) {
  let sibling = direction === "next" ? node.nextSibling : node.previousSibling;
  while (sibling && sibling.nodeType === Node.TEXT_NODE && !(sibling.nodeValue || "").length) {
    sibling = direction === "next" ? sibling.nextSibling : sibling.previousSibling;
  }
  return sibling;
}

function adjacentMathToken(range, direction) {
  const { startContainer, startOffset } = range;

  if (startContainer.nodeType === Node.TEXT_NODE) {
    const text = startContainer.nodeValue || "";
    const boundaryOffset = direction === "next" ? text.length : 0;
    if (startOffset !== boundaryOffset) {
      return null;
    }

    const sibling = adjacentSignificantSibling(startContainer, direction);
    return isMathTokenNode(sibling) ? sibling : null;
  }

  if (startContainer.nodeType === Node.ELEMENT_NODE) {
    const element = startContainer;
    const sibling = direction === "next"
      ? element.childNodes[startOffset]
      : (startOffset > 0 ? element.childNodes[startOffset - 1] : null);
    return isMathTokenNode(sibling) ? sibling : null;
  }

  return null;
}

function nearestEditorChild(node) {
  let current = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
  while (current && current.parentElement && current.parentElement !== editor) {
    current = current.parentElement;
  }
  return current && current.parentElement === editor ? current : null;
}

function nextEditorChild(node) {
  let sibling = node.nextSibling;
  while (sibling && sibling.nodeType === Node.TEXT_NODE && !(sibling.nodeValue || "").trim()) {
    sibling = sibling.nextSibling;
  }
  return sibling;
}

function previousEditorChild(node) {
  let sibling = node.previousSibling;
  while (sibling && sibling.nodeType === Node.TEXT_NODE && !(sibling.nodeValue || "").trim()) {
    sibling = sibling.previousSibling;
  }
  return sibling;
}

function isRangeAtContainerStart(range, container) {
  const probe = document.createRange();
  probe.selectNodeContents(container);
  probe.collapse(true);
  return range.compareBoundaryPoints(Range.START_TO_START, probe) === 0;
}

function isRangeAtContainerEnd(range, container) {
  const probe = document.createRange();
  probe.selectNodeContents(container);
  probe.collapse(false);
  return range.compareBoundaryPoints(Range.END_TO_END, probe) === 0;
}

function mergeWithNextEditorBlock() {
  const range = getCollapsedEditorRange();
  if (!range) {
    return false;
  }

  const current = nearestEditorChild(range.startContainer);
  if (!current || !isRangeAtContainerEnd(range, current)) {
    return false;
  }

  const next = nextEditorChild(current);
  if (!(next instanceof HTMLElement)) {
    return false;
  }

  const firstMoved = next.firstChild;
  if (!firstMoved) {
    next.remove();
    return false;
  }

  while (next.firstChild) {
    current.appendChild(next.firstChild);
  }
  next.remove();
  setCaretBeforeReferenceNode(current, firstMoved);
  scheduleSave();
  return true;
}

function mergeWithPreviousEditorBlock() {
  const range = getCollapsedEditorRange();
  if (!range) {
    return false;
  }

  const current = nearestEditorChild(range.startContainer);
  if (!current || !isRangeAtContainerStart(range, current)) {
    return false;
  }

  const previous = previousEditorChild(current);
  if (!(previous instanceof HTMLElement)) {
    return false;
  }

  const previousLast = previous.lastChild;
  while (current.firstChild) {
    previous.appendChild(current.firstChild);
  }
  current.remove();

  if (previousLast) {
    const rangeAfter = document.createRange();
    rangeAfter.setStartAfter(previousLast);
    rangeAfter.collapse(true);
    const selection = window.getSelection();
    if (selection) {
      selection.removeAllRanges();
      selection.addRange(rangeAfter);
    }
  } else {
    setCaretBeforeNode(previous.firstChild || previous);
  }

  scheduleSave();
  return true;
}

function handleArrowAcrossMathToken(event) {
  if (event.key !== "ArrowRight" && event.key !== "ArrowLeft") {
    return;
  }

  const range = getCollapsedEditorRange();
  if (!range) {
    return;
  }
  const direction = event.key === "ArrowRight" ? "next" : "prev";
  const token = adjacentMathToken(range, direction);
  if (!token) {
    return;
  }

  event.preventDefault();
  const fromSide = direction === "next" ? "from-left" : "from-right";
  expandMathTokenForBoundaryEdit(token, fromSide);
}

function expandMathTokenWithOffset(token, offset) {
  const { raw } = mathRawForToken(token);
  const boundedOffset = Math.max(0, Math.min(offset, raw.length));
  const textNode = document.createTextNode(raw);
  token.replaceWith(textNode);

  editor.focus();
  setCaret(textNode, boundedOffset);
  activeMathEdit = { textNode };
  updateMathPreview();
  scheduleSave();
  return textNode;
}

function expandMathTokenForBoundaryEdit(token, fromSide) {
  const { raw, delimiterSize } = mathRawForToken(token);
  const offset = fromSide === "from-left"
    ? delimiterSize
    : raw.length - delimiterSize;
  return expandMathTokenWithOffset(token, offset);
}

function expandMathTokenForEdit(token, clickX) {
  const { raw } = mathRawForToken(token);

  const rect = token.getBoundingClientRect();
  const ratio = rect.width > 0 ? Math.max(0, Math.min(1, (clickX - rect.left) / rect.width)) : 1;
  const offset = Math.round(ratio * raw.length);
  expandMathTokenWithOffset(token, offset);
}

function finalizeMathEditIfNeeded() {
  if (!activeMathEdit) {
    return;
  }

  const selection = window.getSelection();
  const textNode = activeMathEdit.textNode;
  const stillInDom = editor.contains(textNode);
  if (!selection || selection.rangeCount === 0 || !stillInDom) {
    activeMathEdit = null;
    runMathPass(true);
    return;
  }

  const range = selection.getRangeAt(0);
  const stillEditing = range.startContainer === textNode;

  if (stillEditing) {
    return;
  }

  activeMathEdit = null;
  runMathPass(true);
}

function deleteCharInTextNode(textNode, offset, direction) {
  const value = textNode.nodeValue || "";
  if (direction === "backward") {
    if (offset <= 0) {
      return;
    }
    textNode.nodeValue = `${value.slice(0, offset - 1)}${value.slice(offset)}`;
    setCaret(textNode, offset - 1);
  } else {
    if (offset >= value.length) {
      return;
    }
    textNode.nodeValue = `${value.slice(0, offset)}${value.slice(offset + 1)}`;
    setCaret(textNode, offset);
  }
  scheduleSave();
  updateMathPreview();
}

function editAdjacentMathTokenWithDelete(key) {
  const range = getCollapsedEditorRange();
  if (!range) {
    return false;
  }

  const direction = key === "Backspace" ? "prev" : "next";
  const token = adjacentMathToken(range, direction);
  if (!token) {
    return false;
  }

  const textNode = expandMathTokenWithOffset(
    token,
    direction === "prev" ? mathRawForToken(token).raw.length : 0
  );
  deleteCharInTextNode(textNode, direction === "prev" ? textNode.nodeValue.length : 0, direction === "prev" ? "backward" : "forward");
  return true;
}

function getActiveMathSnippet() {
  const range = getCollapsedEditorRange();
  if (!range || range.startContainer.nodeType !== Node.TEXT_NODE) {
    return null;
  }

  const textNode = range.startContainer;
  const value = textNode.nodeValue || "";
  if (!value.includes("$")) {
    return null;
  }

  const cursor = range.startOffset;
  const regex = /(\$\$([\s\S]+?)\$\$)|(\$([^$\n]+?)\$)/g;
  let match = regex.exec(value);

  while (match) {
    const start = match.index;
    const end = regex.lastIndex;
    if (cursor >= start && cursor <= end) {
      const display = Boolean(match[1]);
      const tex = (display ? match[2] : match[4] || "").trim();
      return {
        tex,
        display,
        range
      };
    }
    match = regex.exec(value);
  }

  return null;
}

function getRangeClientRect(range) {
  const rects = range.getClientRects();
  if (rects.length > 0) {
    return rects[0];
  }

  const fallback = range.startContainer instanceof Element
    ? range.startContainer.getBoundingClientRect()
    : range.startContainer.parentElement?.getBoundingClientRect();

  return fallback || editor.getBoundingClientRect();
}

function hideMathPreview() {
  mathPreview.hidden = true;
}

function updateMathPreview() {
  const snippet = getActiveMathSnippet();
  if (!snippet || !snippet.tex) {
    hideMathPreview();
    return;
  }

  try {
    mathPreviewContent.innerHTML = katex.renderToString(snippet.tex, {
      throwOnError: false,
      strict: "ignore",
      displayMode: snippet.display
    });
  } catch {
    mathPreviewContent.textContent = snippet.tex;
  }

  mathPreview.hidden = false;

  const caretRect = getRangeClientRect(snippet.range);
  const panelWidth = mathPreview.offsetWidth || 320;
  const panelHeight = mathPreview.offsetHeight || 120;

  let left = Math.max(8, Math.min(caretRect.left, window.innerWidth - panelWidth - 8));
  let top = caretRect.top - panelHeight - 10;

  if (top < 8) {
    top = Math.min(window.innerHeight - panelHeight - 8, caretRect.bottom + 12);
  }

  mathPreview.style.left = `${left}px`;
  mathPreview.style.top = `${top}px`;
}

toolbar.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) {
    return;
  }

  const cmd = target.getAttribute("data-cmd");
  const action = target.getAttribute("data-action");

  editor.focus();

  if (cmd) {
    document.execCommand(cmd, false);
    scheduleSave();
    return;
  }

  if (action === "h1") {
    applyHeading(1);
    scheduleSave();
    return;
  }

  if (action === "h2") {
    applyHeading(2);
    scheduleSave();
    return;
  }

  if (action === "math") {
    insertMathByPrompt();
    return;
  }

  if (action === "export-md") {
    exportMarkdown();
  }
});

editor.addEventListener("input", () => {
  scheduleMathPass(false);
  scheduleSave();
  updateMathPreview();
});

editor.addEventListener("blur", () => {
  runMathPass(true);
  hideMathPreview();
});

editor.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }

  const token = target.closest(".math-token");
  if (!token) {
    updateMathPreview();
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  if (activeMathEdit) {
    activeMathEdit = null;
    runMathPass(true);
  }
  expandMathTokenForEdit(token, event.clientX);
});

editor.addEventListener("keydown", (event) => {
  if ((event.key === "Delete" || event.key === "Backspace") && editAdjacentMathTokenWithDelete(event.key)) {
    event.preventDefault();
    return;
  }

  if (event.key === "Delete" && mergeWithNextEditorBlock()) {
    event.preventDefault();
    queueMicrotask(updateMathPreview);
    return;
  }

  if (event.key === "Backspace" && mergeWithPreviousEditorBlock()) {
    event.preventDefault();
    queueMicrotask(updateMathPreview);
    return;
  }

  handleArrowAcrossMathToken(event);
  queueMicrotask(updateMathPreview);
});

editor.addEventListener("mouseup", () => {
  updateMathPreview();
});

importInput.addEventListener("change", async () => {
  const [file] = importInput.files || [];
  if (!file) {
    return;
  }

  const text = await file.text();
  const html = sanitizeImportedHtml(md.render(text));
  editor.innerHTML = html;
  runMathPass(true);
  scheduleSave();

  importInput.value = "";
});

window.addEventListener("keydown", (event) => {
  const key = event.key.toLowerCase();
  if ((event.metaKey || event.ctrlKey) && key === "s") {
    event.preventDefault();
    exportMarkdown();
    return;
  }

  if ((event.metaKey || event.ctrlKey) && key === "m") {
    event.preventDefault();
    insertMathByPrompt();
  }
});

document.addEventListener("selectionchange", () => {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    hideMathPreview();
    return;
  }

  const anchor = selection.anchorNode;
  if (!anchor || !editor.contains(anchor)) {
    finalizeMathEditIfNeeded();
    hideMathPreview();
    return;
  }

  finalizeMathEditIfNeeded();
  updateMathPreview();
});

window.addEventListener("resize", () => {
  updateMathPreview();
});

editor.addEventListener("scroll", () => {
  updateMathPreview();
});
