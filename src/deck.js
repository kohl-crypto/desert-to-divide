const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
// applyTheme comes from Claude's pptx skill (scripts/apply_theme.js). Point PPTX_APPLY_THEME at it.
const { applyTheme } = require(process.env.PPTX_APPLY_THEME || "./apply_theme.js");
const T = require("./data.js");

const THEME = {
  name: "Desert to Divide Pastel",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "3D3A4B", lt1: "FFFFFF", dk2: "6A4FA3", lt2: "F5F0FA",
    accent1: "F7B89A", accent2: "F2A7B8", accent3: "C3B1E1", accent4: "A8D8C9",
    accent5: "B8D8F0", accent6: "7A7489", hlink: "6A4FA3", folHlink: "A63D5E",
  },
};
const INK = { sun: "B4532F", mon: "A63D5E", tue: "6A4FA3", wed: "2F7A66" };
const IMG = (n) => __dirname + "/img/" + n + ".png";

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Desert to Divide: Mari's Colorado Hot Springs Tour Guide";
  pres.author = "Lance";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const dayFill = { sun: C.accent1, mon: C.accent2, tue: C.accent3, wed: C.accent4 };
  const ink = "3D3A4B";
  const I = {};
  for (const [k, comp] of Object.entries({ hike: fa.FaHiking, soak: fa.FaHotTub || fa.FaTint, fun: fa.FaStar, food: fa.FaUtensils, drive: fa.FaCarSide, home: fa.FaFlagCheckered, bed: fa.FaBed, drum: fa.FaDrumstickBite, check: fa.FaCheck, warn: fa.FaExclamationTriangle, map: fa.FaMapMarkedAlt, camera: fa.FaCameraRetro, gas: fa.FaGasPump }))
    I[k] = await icon(comp, ink);

  pres.defineSlideMaster({
    title: "Hero",
    background: { color: C.background2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.95, y: 1.55, w: 5.0, h: 1.2, fontFace: THEME.headFontFace, fontSize: 48, bold: true, color: C.text1, align: "left", valign: "bottom", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.95, y: 2.9, w: 4.9, h: 1.0, fontSize: 16, color: C.accent6, align: "left", valign: "top", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "Day Divider",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.25, w: 5.3, h: 1.3, fontFace: THEME.headFontFace, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "top", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "Title Only",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.75, fontFace: THEME.headFontFace, fontSize: 30, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 9.2, y: 5.2, w: 0.4, h: 0.25, fontSize: 10, color: C.accent6, align: "right" },
  });

  const card = (slide, x, y, w, h, fill, name) => slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.14, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: name,
    shadow: { type: "outer", color: "6A4FA3", opacity: 0.12, blur: 8, offset: 2, angle: 90 },
  });
  const circleIcon = (slide, x, y, d, fill, img, name) => {
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: name + " circle" });
    const p = d * 0.26;
    slide.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: name + " icon", altText: name });
  };
  const text = (slide, t, o) => slide.addText(t, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top" }, o));
  const img = (slide, name, o) => slide.addImage(Object.assign({ path: IMG(name), objectName: "Illustration " + name, altText: "Pastel illustration: " + name }, o));

  // ---------- 1. Title ----------
  pres.addSection({ title: "Welcome" });
  let s = pres.addSlide({ masterName: "Hero", sectionTitle: "Welcome" });
  img(s, "mountains", { x: 0, y: 0, w: 10, h: 5.625 });
  card(s, 0.55, 0.75, 5.8, 4.1, C.background1, "Title card");
  s.addImage({ path: IMG("estonia"), x: 0.95, y: 1.1, w: 0.42, h: 0.27, objectName: "Estonian flag", altText: "Estonian flag" });
  text(s, "FOR MARI  ·  NOV 8–11, 2026", { x: 1.5, y: 1.08, w: 4.3, h: 0.3, fontSize: 12, bold: true, color: INK.tue, charSpacing: 3, valign: "middle" });
  s.addText("Desert to Divide", { placeholder: "title" });
  s.addText("Colorado Springs → Buena Vista → Gunnison → Ouray → Salida → Woodland Park", { placeholder: "body" });
  text(s, "Four days · five hot springs · routed around the flooded roads", { x: 0.95, y: 4.15, w: 5, h: 0.35, fontSize: 13, italic: true, color: INK.mon });
  s.addNotes("Tere, Mari! Four days, five hot springs, and a route that avoids the flooded roads around Pagosa Springs and Durango. The interactive app has a countdown, check-offs, map buttons and the gobble button.");

  // ---------- 2. At a glance ----------
  s = pres.addSlide({ masterName: "Title Only", sectionTitle: "Welcome" });
  s.addText("The trip at a glance", { placeholder: "title" });
  const statFill = [C.accent1, C.accent2, C.accent3, C.accent4];
  T.stats.forEach((st, i) => {
    const x = 0.5 + i * 2.27;
    card(s, x, 1.2, 2.07, 0.95, statFill[i], "Stat " + st.l);
    text(s, st.n, { x, y: 1.24, w: 2.07, h: 0.55, fontFace: THEME.headFontFace, fontSize: 30, bold: true, align: "center", valign: "middle" });
    text(s, st.l.toUpperCase(), { x, y: 1.8, w: 2.07, h: 0.25, fontSize: 11, bold: true, align: "center", charSpacing: 2 });
  });
  T.days.forEach((d, i) => {
    const x = 0.5 + i * 2.27, y = 2.4;
    card(s, x, y, 2.07, 2.75, C.background1, "Day card " + d.day);
    img(s, d.art, { x: x + 0.12, y: y + 0.12, w: 1.83, h: 1.03 });
    text(s, `${d.day.toUpperCase()} · ${d.date}`, { x: x + 0.15, y: y + 1.27, w: 1.8, h: 0.22, fontSize: 10, bold: true, color: INK[d.id] });
    text(s, d.title, { x: x + 0.15, y: y + 1.5, w: 1.8, h: 0.7, fontFace: THEME.headFontFace, fontSize: 13, bold: true, fit: "shrink" });
    text(s, d.id === "wed" ? "Home by 2 PM" : "Sleep: " + d.sleep.town.split(" /")[0], { x: x + 0.15, y: y + 2.33, w: 1.8, h: 0.25, fontSize: 11, color: C.accent6 });
  });

  // ---------- 3. Route ----------
  s = pres.addSlide({ masterName: "Title Only", sectionTitle: "Welcome" });
  s.addText("The route", { placeholder: "title" });
  const nodes = [
    ["Colorado Springs", "Land + rental car", "sun"], ["Woodland Park", "Dinner + gobble", "sun"], ["Cottonwood", "Starlit soak · Sleep", "sun"],
    ["Mt Princeton", "Chalk Creek soak", "mon"], ["Monarch Pass", "Continental Divide", "mon"], ["Ouray", "Pool at dusk · Sleep", "mon"],
    ["Orvis, Ridgway", "Morning soak", "tue"], ["Blue Mesa", "US-50 to Gunnison", "tue"], ["Salida", "Sunset soak · Sleep", "tue"],
    ["Salida", "Bonus sunrise soak", "wed"], ["Wilkerson Pass", "Pikes Peak view", "wed"], ["Charis Bible College", "Finish by 2 PM", "wed"],
  ];
  const lineHex = { sun: "F7B89A", mon: "F2A7B8", tue: "C3B1E1", wed: "A8D8C9" };
  const pos = (i) => { const r = Math.floor(i / 4), c = i % 4; const col = r % 2 === 0 ? c : 3 - c; return { x: 0.9 + col * 2.35, y: 1.25 + r * 1.3 }; };
  for (let i = 0; i < nodes.length - 1; i++) {
    const a = pos(i), b = pos(i + 1);
    const x1 = a.x + 0.17, y1 = a.y + 0.17, x2 = b.x + 0.17, y2 = b.y + 0.17;
    s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001), flipH: x2 < x1, line: { color: lineHex[nodes[i + 1][2]], width: 4, dashType: "dash" }, objectName: "Route segment " + (i + 1) });
  }
  nodes.forEach(([n, sub, d], i) => {
    const p = pos(i);
    s.addShape(pres.shapes.OVAL, { x: p.x, y: p.y, w: 0.34, h: 0.34, fill: { color: dayFill[d] }, line: { color: C.background1, width: 2 }, objectName: "Stop " + n });
    text(s, String(i + 1), { x: p.x, y: p.y, w: 0.34, h: 0.34, fontSize: 10, bold: true, align: "center", valign: "middle" });
    text(s, [{ text: n, options: { bold: true, fontSize: 14, breakLine: true } }, { text: sub, options: { fontSize: 11, color: C.accent6 } }], { x: p.x + 0.42, y: p.y + 0.3, w: 1.85, h: 0.6 });
  });
  [["sun", "Sun"], ["mon", "Mon"], ["tue", "Tue"], ["wed", "Wed"]].forEach(([d, l], i) => {
    s.addShape(pres.shapes.OVAL, { x: 0.9 + i * 1.1, y: 5.12, w: 0.16, h: 0.16, fill: { color: dayFill[d] }, line: { color: dayFill[d], width: 0 }, objectName: "Legend " + l });
    text(s, l, { x: 1.12 + i * 1.1, y: 5.06, w: 0.8, h: 0.28, fontSize: 11, color: C.accent6, valign: "middle" });
  });

  // ---------- Day dividers + timelines ----------
  const kindIcon = { hike: I.hike, soak: I.soak, fun: I.fun, food: I.food, drive: I.drive, home: I.home };
  const highlights = {
    sun: ["Dinner in Woodland Park (and the turkey noise)", "Night drive over Wilkerson Pass", "Cottonwood Hot Springs under the new moon"],
    mon: ["Mt Princeton under the Chalk Cliffs", "Monarch Pass: the Continental Divide", "Ouray Hot Springs Pool at dusk"],
    tue: ["Orvis Hot Springs, Ridgway", "US-50 over Blue Mesa", "Sunset soak in Salida"],
    wed: ["A bonus sunrise soak in Salida", "South Park and Wilkerson Pass", "Finish at Charis Bible College by 2 PM"],
  };
  function divider(d) {
    const sl = pres.addSlide({ masterName: "Day Divider", sectionTitle: "Day by day" });
    img(sl, d.art + "-tall", { x: 6.25, y: 0, w: 3.75, h: 5.625 });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 0.6, w: 1.6, h: 0.45, rectRadius: 0.1, fill: { color: dayFill[d.id] }, line: { color: dayFill[d.id], width: 0 }, objectName: "Day chip" });
    text(sl, `${d.day.toUpperCase()} · ${d.date.toUpperCase()}`, { x: 0.6, y: 0.6, w: 1.6, h: 0.45, fontSize: 13, bold: true, align: "center", valign: "middle", charSpacing: 1 });
    sl.addText(d.title, { placeholder: "title" });
    text(sl, d.route, { x: 0.6, y: 2.6, w: 5.3, h: 0.35, fontSize: 14, color: C.accent6 });
    text(sl, "🚗 " + d.drive, { x: 0.6, y: 2.95, w: 5.3, h: 0.3, fontSize: 13, bold: true, color: INK[d.id] });
    highlights[d.id].forEach((h, i) => {
      const y = 3.45 + i * 0.52;
      circleIcon(sl, 0.6, y, 0.38, dayFill[d.id], I.fun, h);
      text(sl, h, { x: 1.1, y: y + 0.02, w: 4.9, h: 0.36, fontSize: 14, valign: "middle" });
    });
    sl.addNotes(d.title + " — " + d.route + ". " + d.drive + ".");
  }
  function timeline(d, side) {
    const sl = pres.addSlide({ masterName: "Title Only", sectionTitle: "Day by day" });
    sl.addText([{ text: `${d.day} ${d.date}  `, options: { color: INK[d.id] } }, { text: "The plan" }], { placeholder: "title" });
    const stops = d.stops, two = stops.length > 6 && !side, perCol = two ? Math.ceil(stops.length / 2) : stops.length;
    const rowH = two ? 0.74 : Math.min(0.66, 4.1 / stops.length);
    stops.forEach((st, i) => {
      const col = two ? Math.floor(i / perCol) : 0, r = two ? i % perCol : i;
      const x = 0.5 + col * 4.6, y = 1.2 + r * rowH, w = two ? 4.4 : (side ? 5.6 : 9);
      circleIcon(sl, x, y + 0.04, 0.44, st.star ? dayFill[d.id] : C.background2, kindIcon[st.kind] || I.fun, st.name);
      text(sl, st.t, { x: x + 0.55, y: y + 0.02, w: 1.2, h: 0.22, fontSize: 10, bold: true, color: INK[d.id] });
      text(sl, st.name, { x: x + 0.55, y: y + 0.22, w: w - 0.6, h: 0.3, fontSize: 14, bold: !!st.star, fit: "shrink" });
    });
    sl.addNotes(stops.map((st) => `${st.t} — ${st.name}: ${st.desc || ""}${st.options ? " Options: " + st.options.join(" | ") : ""}`).join("\n"));
    return sl;
  }
  pres.addSection({ title: "Day by day" });

  // SUNDAY
  divider(T.days[0]);
  s = timeline(T.days[0], true);
  card(s, 6.3, 1.2, 3.2, 3.85, C.background2, "Sunday eats card");
  circleIcon(s, 6.5, 1.37, 0.45, C.accent1, I.food, "Sunday food");
  text(s, "Where to eat", { x: 7.05, y: 1.42, w: 2.3, h: 0.35, fontFace: THEME.headFontFace, fontSize: 18, bold: true, valign: "middle" });
  text(s, T.days[0].food.flatMap((f) => [
    { text: f.meal.toUpperCase(), options: { fontSize: 10, bold: true, color: INK.sun, breakLine: true } },
    { text: "A: " + f.a, options: { fontSize: 12, breakLine: true } },
    { text: "B: " + f.b, options: { fontSize: 12, breakLine: true, paraSpaceAfter: 8 } },
  ]), { x: 6.5, y: 1.98, w: 2.85, h: 2.9, fit: "shrink" });

  // Buc-ee's mission
  s = pres.addSlide({ masterName: "Title Only", sectionTitle: "Day by day" });
  s.background = { color: "FDEEDC" };
  img(s, "turkey-tall", { x: 6.4, y: 0, w: 3.6, h: 5.625 });
  text(s, "MISSION: WOODLAND PARK DINNER", { x: 0.5, y: 0.4, w: 5.6, h: 0.3, fontSize: 12, bold: true, color: INK.mon, charSpacing: 3 });
  text(s, "Thanksgiving is 18 days away. Practice…", { x: 0.5, y: 0.75, w: 5.7, h: 0.95, fontFace: THEME.headFontFace, fontSize: 26, bold: true });
  text(s, "GOBBLE GOBBLE.", { x: 0.5, y: 1.7, w: 5.8, h: 0.75, fontFace: THEME.headFontFace, fontSize: 44, bold: true, color: INK.mon });
  [[I.drum, "Hearty dinner", "Hot and quick on US-24"], [I.gas, "Fill the tank", "Next sure gas: Buena Vista"],
   [I.camera, "Pikes Peak glow", "Last light as you climb Ute Pass"], [I.fun, "Gobble in the app", "Tap the turkey button"]].forEach(([ic, h, sub], i) => {
    const x = 0.5 + (i % 2) * 2.9, y = 2.7 + Math.floor(i / 2) * 1.05;
    card(s, x, y, 2.75, 0.9, C.background1, "Mission " + h);
    circleIcon(s, x + 0.15, y + 0.2, 0.5, [C.accent1, C.accent2, C.accent3, C.accent4][i], ic, h);
    text(s, h, { x: x + 0.78, y: y + 0.14, w: 1.9, h: 0.3, fontSize: 13, bold: true, fit: "shrink" });
    text(s, sub, { x: x + 0.78, y: y + 0.45, w: 1.9, h: 0.35, fontSize: 11, color: C.accent6 });
  });
  text(s, "Woodland Park, CO · ~40 min west of Colorado Springs Airport", { x: 0.5, y: 4.95, w: 5.8, h: 0.3, fontSize: 11, color: C.accent6 });
  s.addNotes("Commit to the turkey noise. Loud. Non-negotiable.");

  divider(T.days[1]); timeline(T.days[1]);
  divider(T.days[2]); timeline(T.days[2]);
  divider(T.days[3]);
  s = timeline(T.days[3], true);
  card(s, 6.3, 1.2, 3.2, 3.85, C.accent4, "Deadline card");
  circleIcon(s, 6.5, 1.37, 0.45, C.background1, I.home, "Deadline");
  text(s, "Hard deadline", { x: 7.05, y: 1.42, w: 2.3, h: 0.35, fontFace: THEME.headFontFace, fontSize: 18, bold: true, valign: "middle" });
  text(s, "2:00 PM", { x: 6.5, y: 2.05, w: 2.85, h: 0.75, fontFace: THEME.headFontFace, fontSize: 44, bold: true });
  text(s, "Charis Bible College, Woodland Park. Leave Salida by 9:45 AM to arrive around 12:45 PM.", { x: 6.5, y: 2.9, w: 2.85, h: 1.0, fontSize: 13 });
  text(s, "Lunch: something quick in Woodland Park.",{ x: 6.5, y: 4.1, w: 2.85, h: 0.7, fontSize: 12, italic: true });

  // ---------- Hikes ----------
  pres.addSection({ title: "Hikes & stops" });
  s = pres.addSlide({ masterName: "Title Only", sectionTitle: "Hikes & stops" });
  s.addText("The big four soaks", { placeholder: "title" });
  const hikes = T.days.flatMap((d) => d.stops.filter((x) => x.kind === "soak" && x.options).map((h) => ({ d, h }))).slice(0, 4);
  const hikeArt = ["mountains", "black", "canyon", "road"];
  hikes.forEach(({ d, h }, i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.15 + Math.floor(i / 2) * 2.05;
    card(s, x, y, 4.4, 1.9, C.background2, "Hike card " + (i + 1));
    img(s, hikeArt[i] + "-sq", { x: x + 0.12, y: y + 0.12, w: 1.25, h: 1.66, sizing: { type: "cover", w: 1.25, h: 1.66 } });
    text(s, `${d.day} · ${h.t}`, { x: x + 1.5, y: y + 0.14, w: 2.8, h: 0.22, fontSize: 10, bold: true, color: INK[d.id] });
    text(s, h.short || h.name.replace(/^SOAK \d: /, ""), { x: x + 1.5, y: y + 0.36, w: 2.8, h: 0.3, fontSize: 14, bold: true, fit: "shrink" });
    text(s, h.options.map((o, j) => ({ text: o.replace(/^([A-C])[^:]*:\s*/, "$1: ").replace(/\. .*$/, ""), options: { breakLine: j < h.options.length - 1, paraSpaceAfter: 3 } })),
      { x: x + 1.5, y: y + 0.72, w: 2.8, h: 1.08, fontSize: 11, fit: "shrink" });
  });
  s.addNotes(hikes.map(({ h }) => `${h.name}: ${h.desc}\n${h.options.join("\n")}`).join("\n\n"));

  // ---------- Quirky ----------
  s = pres.addSlide({ masterName: "Title Only", sectionTitle: "Hikes & stops" });
  s.addText("Quirky, out-of-the-way stops", { placeholder: "title" });
  const quirky = [
    ["sun", "Wilkerson Pass", "A night crossing at 9,507 ft into South Park."],
    ["mon", "Chalk Cliffs", "White cliffs above Mt Princeton's creekside pools."],
    ["mon", "Monarch Pass", "Continental Divide selfie at 11,312 ft."],
    ["mon", "Black Canyon rim", "A 2,000-ft drop, 2 minutes from the car."],
    ["tue", "Orvis, Ridgway", "A quiet, old-school spring under the snowy San Juans."],
    ["tue", "Blue Mesa", "Dillon Pinnacles and mirror-still water."],
    ["tue", "Salida sunset", "A soak while the Collegiates turn pink."],
    ["wed", "South Park", "A huge empty basin with elk and pronghorn."],
  ];
  quirky.forEach(([d, h, sub], i) => {
    const x = 0.5 + (i % 4) * 2.27, y = 1.15 + Math.floor(i / 4) * 2.0;
    card(s, x, y, 2.07, 1.85, C.background1, "Quirky " + h);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.12, y: y + 0.12, w: 1.83, h: 0.5, rectRadius: 0.1, fill: { color: dayFill[d] }, line: { color: dayFill[d], width: 0 }, objectName: "Quirky tint " + h });
    s.addImage({ data: I.fun, x: x + 0.24, y: y + 0.24, w: 0.26, h: 0.26, objectName: "Star " + h, altText: "star" });
    text(s, d.toUpperCase(), { x: x + 0.6, y: y + 0.24, w: 1.2, h: 0.26, fontSize: 11, bold: true, valign: "middle" });
    text(s, h, { x: x + 0.18, y: y + 0.75, w: 1.75, h: 0.35, fontSize: 14, bold: true, fit: "shrink" });
    text(s, sub, { x: x + 0.18, y: y + 1.1, w: 1.75, h: 0.65, fontSize: 11, color: C.accent6 });
  });

  // ---------- Food ----------
  pres.addSection({ title: "Eat & sleep" });
  s = pres.addSlide({ masterName: "Title Only", sectionTitle: "Eat & sleep" });
  s.addText("Food cheat sheet", { placeholder: "title" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.text1, fill: { color: C.accent3 }, fontSize: 11 } });
  const rows = [[hdr("Day"), hdr("Meal"), hdr("Option A"), hdr("Option B")]];
  T.days.forEach((d) => d.food.forEach((f, j) => rows.push([
    { text: j === 0 ? `${d.day} ${d.date}` : "", options: { bold: true, color: INK[d.id] } },
    { text: f.meal, options: { bold: true } }, { text: f.a }, { text: f.b },
  ])));
  s.addTable(rows, { x: 0.5, y: 1.15, w: 9.0, colW: [0.95, 1.25, 3.4, 3.4], fontSize: 10, fontFace: THEME.bodyFontFace, color: C.text1,
    border: { type: "solid", pt: 0.5, color: "ECE6F3" }, fill: { color: C.background1 }, valign: "middle", margin: [2, 5, 2, 5], rowH: 0.3 });

  // ---------- Sleep ----------
  s = pres.addSlide({ masterName: "Title Only", sectionTitle: "Eat & sleep" });
  s.addText("Where to sleep", { placeholder: "title" });
  const sleepArt = { sun: "mountains", mon: "black", tue: "canyon" };
  T.days.filter((d) => d.sleep).forEach((d, i) => {
    const x = 0.5 + i * 3.07, y = 1.1;
    card(s, x, y, 2.87, 4.05, C.background2, "Sleep " + d.sleep.town);
    img(s, sleepArt[d.id], { x: x + 0.12, y: y + 0.12, w: 2.63, h: 0.85, sizing: { type: "cover", w: 2.63, h: 0.85 } });
    text(s, [{ text: `${d.day} ${d.date}`, options: { fontSize: 10, bold: true, color: INK[d.id], breakLine: true } }, { text: d.sleep.town.split(" /")[0], options: { fontSize: 16, bold: true, fontFace: THEME.headFontFace } }], { x: x + 0.18, y: y + 1.07, w: 2.5, h: 0.55 });
    text(s, d.sleep.picks.flatMap((p, j) => [
      { text: (j + 1) + ". “" + p.name + "”", options: { bold: true, fontSize: 11, breakLine: true } },
      { text: p.note, options: { fontSize: 10, color: C.accent6, breakLine: j < 2, paraSpaceAfter: 5 } },
    ]), { x: x + 0.18, y: y + 1.7, w: 2.55, h: 2.25, fit: "shrink" });
  });
  text(s, "These are ideas, not bookings. Check dates, price and parking, then search Airbnb by town to compare.", { x: 0.5, y: 5.2, w: 8.5, h: 0.25, fontSize: 10, color: C.accent6 });

  // ---------- Pack ----------
  pres.addSection({ title: "Before you go" });
  s = pres.addSlide({ masterName: "Title Only", sectionTitle: "Before you go" });
  s.addText("Pack it & know before you go", { placeholder: "title" });
  card(s, 0.5, 1.15, 4.3, 3.95, C.background2, "Packing card");
  circleIcon(s, 0.68, 1.3, 0.45, C.accent4, I.check, "Packing");
  text(s, "Pack list", { x: 1.25, y: 1.35, w: 3, h: 0.35, fontFace: THEME.headFontFace, fontSize: 18, bold: true, valign: "middle" });
  text(s, T.deckPack.map((c, i, a) => ({ text: c, options: { bullet: true, breakLine: i < a.length - 1 } })), { x: 0.7, y: 1.9, w: 3.95, h: 3.1, fontSize: 12, paraSpaceAfter: 3, fit: "shrink" });
  card(s, 5.2, 1.15, 4.3, 3.95, C.background2, "Heads up card");
  circleIcon(s, 5.38, 1.3, 0.45, C.accent2, I.warn, "Heads up");
  text(s, "Heads up", { x: 5.95, y: 1.35, w: 3, h: 0.35, fontFace: THEME.headFontFace, fontSize: 18, bold: true, valign: "middle" });
  text(s, T.deckHeads.map((c, i, a) => ({ text: c, options: { bullet: true, breakLine: i < a.length - 1 } })), { x: 5.4, y: 1.9, w: 3.95, h: 3.1, fontSize: 12, paraSpaceAfter: 3, fit: "shrink" });

  // ---------- Closing ----------
  s = pres.addSlide({ masterName: "Hero", sectionTitle: "Before you go" });
  img(s, "road", { x: 0, y: 0, w: 10, h: 5.625 });
  card(s, 0.55, 0.75, 5.8, 4.1, C.background1, "Closing card");
  s.addImage({ path: IMG("estonia"), x: 0.95, y: 1.1, w: 0.42, h: 0.27, objectName: "Estonian flag", altText: "Estonian flag" });
  text(s, "HEAD REISI, MARI!", { x: 1.5, y: 1.08, w: 4.3, h: 0.3, fontSize: 12, bold: true, color: INK.tue, charSpacing: 3, valign: "middle" });
  s.addText([{ text: "Have the best time!", options: { fontSize: 40 } }], { placeholder: "title" });
  s.addText("Gobble gobble. 🦃 The interactive app has the countdown, check-offs, map buttons, and the gobble button.", { placeholder: "body" });

  const OUT = __dirname + "/../Desert-to-Divide-Tour-Guide.pptx";
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok");
})();
