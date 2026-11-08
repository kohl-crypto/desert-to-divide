const fs=require("fs");
const {ART,ESTONIA}=require(__dirname+"/art.js");
const src=fs.readFileSync(__dirname+"/src.html","utf8");
const data=fs.readFileSync(__dirname+"/data.js","utf8")+"\n"+fs.readFileSync(__dirname+"/art.js","utf8");
const audio=fs.readFileSync(__dirname+"/audio/gobble.mp3").toString("base64");
const dl=`<a class="btn primary" style="--day:#F7B0C4" href="Desert-to-Divide-Tour-Guide.pptx" download>⬇ PowerPoint guide</a><a class="btn" href="index.html" download="Desert-to-Divide.html">⬇ Save app for offline</a>`;
const build=(downloads)=>src.replace("__FLAG__",()=>encodeURIComponent(ESTONIA)).replace("__HEROART__",()=>ART.mountains)
  .replace("__DOWNLOADS__",()=>downloads).replace("__AUDIO__",()=>audio).replace("/*__DATA__*/",()=>data);
fs.writeFileSync(__dirname+"/../index.html",build(dl));


console.log("built index.html", fs.statSync(__dirname+"/../index.html").size);
