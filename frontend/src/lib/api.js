import axios from "axios";

const BACKEND_URL = (process.env.REACT_APP_BACKEND_URL || "http://localhost:8000").replace(/\/$/, "");
export const API = `${BACKEND_URL}/api`;

export const api = axios.create({ baseURL: API, timeout: 120000 });

export const fmt = (v, nd = 1) => {
  if (v === null || v === undefined || v === "") return "-";
  const f = Number(v);
  if (Number.isNaN(f)) return String(v);
  return f.toFixed(nd);
};

export const DOT_COLORS = {
  RED: "#FF0000",
  ORANGE: "#FFC000",
  GREEN: "#00B050",
  NONE: "#BFBFBF",
};

export const VERDICT_COLORS = {
  "Low Risk": "#00B050",
  "Medium Risk": "#FFC000",
  "High Risk": "#FF0000",
};

export const CELL_COLORS = {
  RED: "#FF0000",
  YELLOW: "#FFFF00",
  GREEN: "#00B050",
};

const CHANNEL_ALIASES = {
  "vehicle speed": ["p: vehicle speed, vehicle speed", "vehicle speed"],
  "accelerationchassis": ["accelerationchassis", "val_a_tstart, accelerationchassis"],
  "throttle position": ["val_a_tstart, throttle position", "max, throttle position"],
  "max throttle position": ["max, throttle position", "val_a_tstart, throttle position"],
  "sub event name": ["sous situation de vie, sub event name", "sub event name"],
  "gear new": ["p: gear end, gear end", "p: gear new, gear new", "gear new"],
};

export function getChannel(channels, name) {
  if (!channels || name == null) return null;
  if (channels[name] !== undefined) return channels[name];
  const ln = String(name).replace(/\u00b7/g, ".").trim().toLowerCase();
  const lookup = (target) => {
    for (const [k, v] of Object.entries(channels)) {
      const kk = String(k).replace(/\u00b7/g, ".").trim().toLowerCase();
      if (kk === target) return v;
      if (kk.includes(",") && kk.split(",").slice(1).join(",").trim() === target) return v;
    }
    return undefined;
  };
  let hit = lookup(ln);
  if (hit !== undefined) return hit;
  if (ln.includes(",")) {
    hit = lookup(ln.split(",").slice(1).join(",").trim());
    if (hit !== undefined) return hit;
  }
  for (const alias of CHANNEL_ALIASES[ln] || []) {
    hit = lookup(alias);
    if (hit !== undefined) return hit;
  }
  return null;
}

/** Case-insensitive map get for criterion / target keys. */
export function resolveKey(map, name) {
  if (!map || name == null) return null;
  if (map[name] !== undefined) return map[name];
  const ln = String(name).replace(/\u00b7/g, ".").trim().toLowerCase();
  for (const [k, v] of Object.entries(map)) {
    if (String(k).replace(/\u00b7/g, ".").trim().toLowerCase() === ln) return v;
  }
  return null;
}

export async function downloadBlob(url, body, filename) {
  const res = await api.post(url, body, { responseType: "blob" });
  const blob = new Blob([res.data]);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
}
