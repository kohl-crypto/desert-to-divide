// Pastel scene illustrations (original SVG art) shared by the app and the deck.
const P = {
  sky1: "#FBE3EC", sky2: "#FDEEDC", sky3: "#E9E4F7", sun: "#FFD8A8",
  peach: "#F7B89A", rose: "#F2A7B8", lav: "#C3B1E1", mint: "#A8D8C9",
  blue: "#B8D8F0", sand: "#F6DCC4", clay: "#E8A48C", plum: "#8E7BB5", deep: "#6E5A96", ink: "#3D3A4B",
};
const W = 800, H = 450;
const frame = (id, defs, body) => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid slice" role="img">
<defs><linearGradient id="${id}s" x1="0" y1="0" x2="0" y2="1">${defs}</linearGradient></defs>
<rect width="${W}" height="${H}" fill="url(#${id}s)"/>${body}</svg>`;
const stops = (...c) => c.map((col, i) => `<stop offset="${(i / (c.length - 1)).toFixed(2)}" stop-color="${col}"/>`).join("");
const sun = (x, y, r, c = P.sun) => `<circle cx="${x}" cy="${y}" r="${r * 1.6}" fill="${c}" opacity=".35"/><circle cx="${x}" cy="${y}" r="${r}" fill="${c}"/>`;
const cloud = (x, y, s = 1) => `<g fill="#fff" opacity=".8" transform="translate(${x} ${y}) scale(${s})"><ellipse cx="0" cy="0" rx="46" ry="16"/><ellipse cx="-22" cy="-8" rx="22" ry="16"/><ellipse cx="16" cy="-12" rx="26" ry="20"/></g>`;

const ART = {
  canyon: frame("gc", stops(P.sky3, P.sky1, P.sky2),
    sun(610, 120, 46) + cloud(180, 90, 1.1) + cloud(470, 60, .8) +
    `<path d="M0 210 L120 200 L170 215 L300 205 L360 222 L520 208 L600 220 L800 205 V450 H0Z" fill="${P.lav}"/>
     <path d="M0 250 L90 245 L140 262 L260 250 L330 270 L470 252 L560 268 L680 255 L800 262 V450 H0Z" fill="${P.rose}"/>
     <path d="M0 295 L80 288 L150 305 L230 296 L320 318 L420 300 L520 320 L640 302 L800 312 V450 H0Z" fill="${P.peach}"/>
     <path d="M0 340 L120 332 L210 350 L300 340 L360 365 L440 345 L560 362 L680 346 L800 355 V450 H0Z" fill="${P.clay}"/>
     <path d="M300 450 C340 400 380 410 400 380 C420 410 470 400 510 450Z" fill="${P.blue}"/>
     <path d="M0 395 L140 385 L250 405 L330 450 H0Z M800 395 L640 388 L560 410 L500 450 H800Z" fill="${P.plum}" opacity=".55"/>`),

  monument: frame("mv", stops(P.sky1, P.sky2, P.sand),
    sun(400, 230, 70, "#FFC9A8") + cloud(640, 80, .9) +
    `<path d="M110 330 V190 h18 v-30 h28 v30 h20 V330Z" fill="${P.clay}"/>
     <path d="M150 330 V230 h-14 v-24 h-10 v124Z" fill="${P.clay}"/>
     <path d="M430 330 V200 h22 v-36 h30 v36 h18 V330Z" fill="${P.rose}"/>
     <path d="M520 330 V215 h-12 v-28 h-10 v143Z" fill="${P.rose}"/>
     <path d="M600 330 V240 h120 V330Z" fill="${P.peach}"/>
     <path d="M0 320 H800 V450 H0Z" fill="${P.sand}"/>
     <path d="M360 450 L395 320 H405 L440 450Z" fill="${P.lav}"/>
     <path d="M398 450 L399 330 h2 L402 450Z" fill="#fff" opacity=".9"/>`),

  arch: frame("ar", stops(P.sky3, P.sky1, P.sky2),
    sun(160, 150, 40) +
    `<path d="M430 250 L520 150 L560 170 L620 120 L700 175 L800 160 V300 H430Z" fill="#fff"/>
     <path d="M430 250 L520 150 L545 175 L520 210 L560 190 L620 120 L660 160 L640 200 L700 175 L800 160 V300 H430Z" fill="${P.lav}" opacity=".6"/>
     <path d="M0 330 C150 300 300 315 420 325 C560 335 680 300 800 310 V450 H0Z" fill="${P.peach}"/>
     <path d="M250 330 C250 200 280 150 330 140 C390 132 420 190 420 330 L395 330 C395 220 375 175 335 178 C300 182 282 230 282 330Z" fill="${P.clay}"/>
     <path d="M0 380 C200 360 380 372 520 380 C640 388 720 370 800 375 V450 H0Z" fill="${P.rose}"/>`),

  black: frame("bc", stops(P.sky3, P.blue),
    cloud(160, 70, .9) + cloud(620, 100, 1.1) +
    `<path d="M0 140 H330 L360 450 H0Z" fill="${P.plum}"/>
     <path d="M800 150 H450 L420 450 H800Z" fill="${P.deep}"/>
     <path d="M0 140 H330 L335 190 L0 200Z M800 150 H450 L446 195 L800 205Z" fill="${P.mint}"/>
     <path d="M470 220 l-6 60 M480 260 l-8 70 M500 240 l-4 90" stroke="${P.rose}" stroke-width="5" stroke-linecap="round" opacity=".8"/>
     <path d="M362 450 C370 420 410 420 420 450Z" fill="${P.blue}"/>
     <path d="M0 140 H330 M450 150 H800" stroke="#fff" stroke-width="3" opacity=".6"/>`),

  mountains: frame("mn", stops(P.sky3, P.sky1),
    sun(640, 110, 36) +
    `<path d="M0 300 L140 130 L230 230 L330 90 L450 240 L560 150 L700 260 L800 200 V450 H0Z" fill="${P.lav}"/>
     <path d="M140 130 L170 165 L150 160 L125 148Z M330 90 L370 140 L340 128 L310 118Z M560 150 L590 185 L565 175 L540 170Z" fill="#fff"/>
     <path d="M0 340 L180 250 L300 320 L420 230 L560 330 L680 260 L800 320 V450 H0Z" fill="${P.rose}"/>
     <path d="M0 400 L260 330 L420 390 L600 340 L800 395 V450 H0Z" fill="${P.mint}"/>
     <path d="M640 450 C600 430 690 410 650 395 C610 380 700 365 660 350" stroke="#fff" stroke-width="10" fill="none" stroke-linecap="round" opacity=".9"/>`),

  turkey: frame("tk", stops(P.sky2, P.sky1),
    `<g transform="translate(400 250)">
      ${[-75, -50, -25, 0, 25, 50, 75].map((a, i) => `<ellipse cx="0" cy="-95" rx="34" ry="95" fill="${[P.rose, P.peach, P.lav, P.mint, P.lav, P.peach, P.rose][i]}" transform="rotate(${a})"/>`).join("")}
      <ellipse cx="0" cy="40" rx="78" ry="70" fill="${P.clay}"/>
      <circle cx="0" cy="-35" r="40" fill="${P.sand}"/>
      <circle cx="-14" cy="-44" r="6" fill="${P.ink}"/><circle cx="14" cy="-44" r="6" fill="${P.ink}"/>
      <path d="M-10 -28 L10 -28 L0 -12Z" fill="#FFC36B"/>
      <path d="M8 -20 C22 -10 14 8 4 4" fill="#F28B9B"/>
      <path d="M-30 105 v24 M30 105 v24" stroke="#FFC36B" stroke-width="8" stroke-linecap="round"/>
    </g>
    <text x="400" y="420" text-anchor="middle" font-family="Arial, sans-serif" font-weight="900" font-size="40" fill="${P.ink}">GOBBLE GOBBLE!</text>`),

  road: frame("rd", stops(P.sky1, P.sky2),
    sun(400, 200, 60) +
    `<path d="M0 260 L120 240 L180 250 L260 225 L320 245 L480 235 L560 250 L660 230 L800 245 V450 H0Z" fill="${P.lav}"/>
     <path d="M0 300 H800 V450 H0Z" fill="${P.sand}"/>
     <path d="M330 450 L392 300 H408 L470 450Z" fill="${P.plum}" opacity=".55"/>
     <path d="M398 450 L399.4 300 h1.2 L402 450Z" fill="#fff" stroke="none" stroke-dasharray="14 10"/>
     <path d="M120 300 v-50 M110 270 h20 M690 300 v-60 M680 262 h20" stroke="${P.mint}" stroke-width="10" stroke-linecap="round"/>`),
};

const ESTONIA = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 33 21"><rect width="33" height="7" fill="#0072CE"/><rect y="7" width="33" height="7" fill="#000"/><rect y="14" width="33" height="7" fill="#fff"/></svg>`;

if (typeof module !== "undefined") module.exports = { ART, P, ESTONIA };
