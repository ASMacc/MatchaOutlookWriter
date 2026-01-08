const FLOW_ID = "6c1ab2f5-1317-48af-9509-3dafcac71e17";
const MATCHA_BASE = "https://matchaflow.harriscomputer.com";

// Demo-only. For a real deployment, do not store keys client-side.
let apiKey = "";

Office.onReady(() => {
  document.getElementById("btnRewrite").addEventListener("click", onRewrite);
});

function setStatus(msg) {
  document.getElementById("status").textContent = msg;
}

function getBodyText() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    item.body.getAsync("text", (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value || "");
      else reject(res.error);
    });
  });
}

function insertText(text) {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    // Inserts at cursor/selection. Safer than overwriting entire body.
    item.body.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (res) => (res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error))
    );
  });
}

async function runMatcha(inputValue) {
  const url = `${MATCHA_BASE}/api/v1/run/${FLOW_ID}`;

  const body = {
    input_value: inputValue,
    input_type: "chat",
    output_type: "chat",
  };

  const resp = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const t = await resp.text();
    throw new Error(`Matcha error ${resp.status}: ${t}`);
  }

  const r = await resp.json();

  // Your known-good extraction:
  const out0 = r?.outputs?.[0]?.outputs?.[0];
  return (
    out0?.outputs?.text?.message ??
    out0?.messages?.[0]?.message ??
    ""
  );
}

async function onRewrite() {
  try {
    if (!apiKey) {
      apiKey = window.prompt("Enter Langflow x-api-key (demo only):") || "";
      if (!apiKey) return;
    }

    setStatus("Reading email body...");
    const draft = await getBodyText();

    setStatus("Calling Matcha...");
    // Your system prompt expects: body contains instructions inside ```...```
    // So we send the raw draft.
    const result = await runMatcha(draft);

    if (!result) throw new Error("Empty response from flow.");

    setStatus("Inserting result at cursor...");
    await insertText(result);

    setStatus("Done.");
  } catch (e) {
    console.error(e);
    setStatus(`Error: ${e.message || String(e)}`);
  }
}
