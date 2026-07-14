let editor = null;
let rendered = false;
let timer = null;
let lastSentValue = null;

function sendValue(value, wait) {
  window.clearTimeout(timer);
  timer = window.setTimeout(() => {
    lastSentValue = value;
    Streamlit.setComponentValue(value);
  }, Math.max(0, wait || 0));
}

function makeEditor(args) {
  const wrap = document.getElementById("input_wrap");
  wrap.innerHTML = "";
  editor = document.createElement(args.multiline ? "textarea" : "input");
  editor.id = "editor";
  editor.name = "editor";
  if (!args.multiline) editor.type = "text";
  editor.value = args.value || "";
  editor.placeholder = args.placeholder || "";
  editor.disabled = Boolean(args.disabled);
  if (args.max_chars) editor.maxLength = args.max_chars;
  if (args.multiline) {
    editor.style.height = `${Math.max(48, args.input_height || 110)}px`;
  }
  editor.addEventListener("input", (event) => sendValue(event.target.value, args.debounce));
  wrap.appendChild(editor);
  lastSentValue = editor.value;
}

function onRender(event) {
  const args = event.detail.args;
  const theme = event.detail.theme || {};
  const root = document.getElementById("root");
  root.style.setProperty("--primary-color", theme.primaryColor || "#ff4b4b");
  root.style.setProperty("--background-color", theme.backgroundColor || "#ffffff");
  root.style.setProperty("--secondary-background-color", theme.secondaryBackgroundColor || "#f0f2f6");
  root.style.setProperty("--text-color", theme.textColor || "#31333f");
  root.style.setProperty("--font", theme.font || "sans-serif");

  const label = document.getElementById("label");
  label.textContent = args.label || "";
  root.classList.toggle("label-hidden", args.label_visibility === "hidden");
  root.classList.toggle("label-collapsed", args.label_visibility === "collapsed");

  const requestedKind = args.multiline ? "TEXTAREA" : "INPUT";
  if (!rendered || !editor || editor.tagName !== requestedKind) {
    makeEditor(args);
    rendered = true;
  } else {
    editor.placeholder = args.placeholder || "";
    editor.disabled = Boolean(args.disabled);
    if (args.max_chars) editor.maxLength = args.max_chars;
    else editor.removeAttribute("maxlength");
    if (args.multiline) editor.style.height = `${Math.max(48, args.input_height || 110)}px`;

    // Do not overwrite the user's caret while they are typing. External value
    // changes (new upload / reset / other editor) are applied when not focused.
    const incoming = args.value || "";
    if (document.activeElement !== editor && incoming !== editor.value && incoming !== lastSentValue) {
      editor.value = incoming;
      lastSentValue = incoming;
    }
  }

  Streamlit.setFrameHeight(args.frame_height || (args.multiline ? 140 : 73));
}

Streamlit.events.addEventListener(Streamlit.RENDER_EVENT, onRender);
Streamlit.setComponentReady();
