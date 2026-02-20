try {
  // get all parameters from search and hash, combine them into a single string
  const source = `${window.location.search.replace(/^\?/, "") || ""}&${window.location.hash.replace(/^#/, "") || ""}`;
  // get state to decode library info
  const state = new URLSearchParams(source).get("state") || '';
  // base64 decode state and parse library info
  const normalized = state.split("|")[0].replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized + "=".repeat((4 - (normalized.length % 4)) % 4);
  const bytes = Uint8Array.from(atob(padded), (char) => char.charCodeAt(0));
  // get library from state and post message to channel
  const library = JSON.parse(new TextDecoder().decode(bytes));
  // post message to channel with library id and source
  const channel = new BroadcastChannel(library.id);
  channel.postMessage({ v: 1, payload: source });
  // close channel and window
  channel.close();
  window.close();
} catch (e) {}