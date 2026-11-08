const fs=require("fs");
const {ART,ESTONIA}=require(__dirname+"/art.js");
const src=fs.readFileSync(__dirname+"/src.html","utf8");
const data=fs.readFileSync(__dirname+"/data.js","utf8")+"\n"+fs.readFileSync(__dirname+"/art.js","utf8");
const audio=fs.readFileSync(__dirname+"/audio/gobble.mp3").toString("base64");
const anthem=fs.readFileSync(__dirname+"/audio/anthem.mp3").toString("base64");
const dl=`<button class="btn primary installBtn" type="button" style="--day:#A8D8C9">📲 Install app on my phone</button><a class="btn primary" style="--day:#F7B0C4" href="Desert-to-Divide-Tour-Guide.pptx">⬇ PowerPoint guide</a>`;
const build=(downloads)=>src.replace("__FLAG__",()=>encodeURIComponent(ESTONIA)).replace("__HEROART__",()=>ART.mountains)
  .replace("__DOWNLOADS__",()=>downloads).replace("__AUDIO__",()=>audio).replace("__ANTHEM__",()=>anthem).replace("__SLIDES__",()=>String(fs.readdirSync(__dirname+"/../slides").filter(f=>/\.jpg$/.test(f)).length)).replace("__FEEDBACK_URL__",()=>"https://desert-to-divide-relay.vercel.app/api/feedback").replace("__VERSION__",()=>new Date().toISOString().slice(0,16)).replace("/*__DATA__*/",()=>data);
fs.writeFileSync(__dirname+"/../index.html",build(dl));


console.log("built index.html", fs.statSync(__dirname+"/../index.html").size);
