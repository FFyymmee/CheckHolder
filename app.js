var TOKEN="0x2f5c60bdE7A5ebD2b116BB03cB5232fA1Ea55F1C";
var BASE="https://scan.tokenx.finance/api/v2/tokens/"+TOKEN+"/holders";
var rows=[],view=[],page=0,PS=50,sk="balance",sd=-1,info={symbol:"",supply:0,total:0};
function $(i){return document.getElementById(i)}
function say(t,e){var m=$("msg");m.textContent=t;m.className="msg"+(e?" err":"")}
function fmt(n){return isNaN(n)?"-":n.toLocaleString("en-US",{maximumFractionDigits:4})}
function sleep(ms){return new Promise(function(r){setTimeout(r,ms)})}
async function getJSON(url){
  for(var t=0;t<4;t++){
    try{
      var r=await fetch(url);
      if(r.status===429){await sleep(1500*(t+1));continue}
      if(!r.ok)throw new Error("HTTP "+r.status);
      return await r.json();
    }catch(e){if(t===3)throw e;await sleep(800*(t+1))}
  }
}
$("go").onclick=async function(){
  var btn=$("go");btn.disabled=true;rows=[];var seenAddr={};var params="",n=0,guard=0;
  try{
    while(guard++<500){
      var j=await getJSON(BASE+(params?"?"+params:""));
      (j.items||[]).forEach(function(it){
        var d=Number(it.token.decimals||0);
        if(!info.symbol){
          info.symbol=it.token.symbol||"";info.name=it.token.name||"";
          info.supply=Number(it.token.total_supply)/Math.pow(10,d);
          info.total=Number(it.token.holders||0);
          $("title").textContent=(info.name||"Token")+" ("+info.symbol+") Holders";
          $("sub").textContent=TOKEN;
        }
        var key=it.address.hash.toLowerCase();if(seenAddr[key])return;seenAddr[key]=1;
        var bal=Number(it.value)/Math.pow(10,d);
        rows.push({voter:"",grp:"",note:"",address:it.address.hash,type:it.address.is_contract?"Contract":"Wallet",balance:bal,pct:info.supply>0?bal/info.supply*100:NaN});
      });
      n=rows.length;
      $("pb").style.width=(info.total?Math.min(100,n/info.total*100):50)+"%";
      say("โหลดแล้ว "+n.toLocaleString()+(info.total?" / "+info.total.toLocaleString():""));
      if(!j.next_page_params)break;
      params=new URLSearchParams(j.next_page_params).toString();
      await sleep(150);
    }
    rows.sort(function(a,b){return b.balance-a.balance});
    rows.forEach(function(x,i){x.rank=i+1});
    $("pb").style.width="100%";say("โหลดครบ "+rows.length.toLocaleString()+" รายการ");
    sk="balance";sd=-1;page=0;$("out").hidden=false;stats();if(sheetObjs)doMatch(true);apply();
  }catch(e){
    say("โหลดไม่สำเร็จ: "+e.message+(rows.length?" (ได้ "+rows.length+" รายการ)":"")+" — หน้านี้ต้องเปิดจากไฟล์ในเครื่องหรือเว็บของคุณเอง",true);
    if(rows.length){rows.forEach(function(x,i){x.rank=i+1});$("out").hidden=false;stats();apply()}
  }
  btn.disabled=false;btn.textContent="โหลดใหม่";
};
function stats(){
  var tot=rows.reduce(function(s,x){return s+x.balance},0);
  var top10=rows.slice(0,10).reduce(function(s,x){return s+x.balance},0);
  var s=[["จำนวนผู้ถือ",rows.length.toLocaleString()],["รวมที่ผู้ถือถืออยู่",fmt(tot)],["10 อันดับแรกถือ",info.supply>0?(top10/info.supply*100).toFixed(2)+"%":"-"],["Total supply",fmt(info.supply)]];
  $("stats").innerHTML=s.map(function(a){return '<div class="stat"><b>'+a[1]+'</b><span>'+a[0]+'</span></div>'}).join("");
}
function apply(){
  var q=$("q").value.trim().toLowerCase(),f=$("flt").value;
  view=rows.filter(function(x){
    if(f==="yes"&&!x.voter)return false;if(f==="no"&&x.voter)return false;
    return !q||x.address.toLowerCase().indexOf(q)>-1||(x.voter+" "+x.note).toLowerCase().indexOf(q)>-1});
  view.sort(function(a,b){
    var x=a[sk],y=b[sk];
    if(typeof x==="string")return sd*(x<y?-1:x>y?1:0);
    return sd*((x||0)-(y||0));
  });
  render();
}
function render(){
  var pages=Math.max(1,Math.ceil(view.length/PS));if(page>=pages)page=pages-1;
  $("tb").innerHTML=view.slice(page*PS,page*PS+PS).map(function(x){
    return '<tr><td class="n">'+x.rank+'</td><td class="a">'+x.address+'</td><td>'+x.type+'</td><td class="v">'+esc(x.voter)+'</td><td>'+esc(x.note)+'</td><td class="n">'+fmt(x.balance)+'</td><td class="n">'+(isNaN(x.pct)?"-":x.pct.toFixed(4)+"%")+'</td></tr>';
  }).join("");
  $("pi").textContent="หน้า "+(page+1)+" / "+pages+" ("+view.length.toLocaleString()+")";
}
document.querySelectorAll("th").forEach(function(th){th.onclick=function(){
  var k=th.dataset.k;if(sk===k)sd=-sd;else{sk=k;sd=(k==="address"||k==="type"||k==="voter"||k==="grp"||k==="note")?1:-1}apply()}});
$("q").oninput=function(){page=0;apply()};
$("prev").onclick=function(){if(page>0){page--;render()}};
$("next").onclick=function(){page++;render()};
function data(){return rows.map(function(x){return {"Rank":x.rank,"Address":x.address,"Type":x.type,"Voter name":x.voter,"Wallet note":x.note,"Balance":x.balance,"Percent of supply":isNaN(x.pct)?"":Number(x.pct.toFixed(6))}})}
$("xl").onclick=function(){
  var wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data()),"Holders");
  XLSX.writeFile(wb,(info.symbol||"token")+"-holders.xlsx");
};
$("csv").onclick=function(){
  var csv=XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(data()));
  var a=document.createElement("a");a.href=URL.createObjectURL(new Blob(["\ufeff"+csv],{type:"text/csv"}));
  a.download=(info.symbol||"token")+"-holders.csv";a.click();
};

$("flt").onchange=function(){page=0;apply()};
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]})}
var ADDR=/0x[0-9a-fA-F]{40}/;
var sheetWb=null,sheetObjs=null,ge6=null;
function smsg(t,e){var m=$("smsg");m.textContent=t;m.className="msg"+(e?" err":"")}
function openSheetWb(wb){
  sheetWb=wb;ge6=null;
  var sheets=wb.SheetNames.map(function(n){return {name:n,aoa:XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:""})}});
  var ent=parseGE6(sheets);
  if(ent.length){
    ge6=ent;sheetObjs=ent;
    var mem={},ad={};ent.forEach(function(x){mem[x.group+"|"+x.name]=1;ad[x.address]=1});
    $("smap").hidden=true;
    smsg("อ่านไฟล์รวมกระเป๋าแล้ว: "+Object.keys(ad).length.toLocaleString()+" address จากสมาชิก "+Object.keys(mem).length+" รายการ ใน "+sheets.length+" แท็บ — กด จับคู่ ได้เลย");
    $("smatch").hidden=false;$("smap").hidden=false;
    ["stab","scaddr","scname"].forEach(function(i){$(i).parentNode.hidden=true});
    return;
  }
  ["stab","scaddr","scname"].forEach(function(i){$(i).parentNode.hidden=false});
  $("stab").innerHTML=wb.SheetNames.map(function(n,i){return '<option value="'+i+'">'+esc(n)+'</option>'}).join("");
  $("smap").hidden=false;useTab();
}
function useTab(){
  var name=sheetWb.SheetNames[Number($("stab").value)||0];
  sheetObjs=XLSX.utils.sheet_to_json(sheetWb.Sheets[name],{defval:""});
  if(!sheetObjs.length){smsg("แท็บนี้ไม่มีข้อมูล",true);return}
  var heads=Object.keys(sheetObjs[0]),best="",bc=-1;
  heads.forEach(function(h){var c=sheetObjs.filter(function(r){return ADDR.test(String(r[h]))}).length;if(c>bc){bc=c;best=h}});
  var opts=heads.map(function(h){return '<option value="'+esc(h)+'">'+esc(h)+'</option>'}).join("");
  $("scaddr").innerHTML=opts;$("scname").innerHTML=opts;
  if(bc>0)$("scaddr").value=best;
  var nm=heads.filter(function(h){return h!==best&&/ชื่อ|name|สมาชิก|member|user/i.test(h)})[0]||heads.filter(function(h){return h!==best})[0]||best;
  $("scname").value=nm;
  smsg("พบ "+sheetObjs.length.toLocaleString()+" แถว ตรวจคอลัมน์ให้ถูกต้องแล้วกด จับคู่"+(bc<=0?" (ไม่พบคอลัมน์ที่เป็น address อัตโนมัติ กรุณาเลือกเอง)":""));
}
$("stab").onchange=useTab;
$("sfile").onchange=function(e){
  var f=e.target.files[0];if(!f)return;var r=new FileReader();
  r.onload=function(){try{openSheetWb(XLSX.read(r.result,{type:"array"}))}catch(x){smsg("อ่านไฟล์ไม่ได้: "+x.message,true)}};
  r.readAsArrayBuffer(f);
};
$("sget").onclick=async function(){
  var m=$("surl").value.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if(!m){smsg("ลิงก์ไม่ถูกต้อง",true);return}
  var g=$("surl").value.match(/gid=(\d+)/);
  var url="https://docs.google.com/spreadsheets/d/"+m[1]+"/export?format=csv"+(g?"&gid="+g[1]:"");
  try{
    var r=await fetch(url);if(!r.ok)throw new Error("HTTP "+r.status);
    var t=await r.text();if(/^\s*<(!doctype|html)/i.test(t))throw new Error("ชีตไม่เปิดสาธารณะ");
    openSheetWb(XLSX.read(t,{type:"string"}));
  }catch(x){smsg("ดึงจากลิงก์ไม่ได้ ("+x.message+") ให้ดาวน์โหลดไฟล์จากชีตแล้วเลือกไฟล์แทน",true)}
};
function doMatchGE6(quiet){
  var map={};
  ge6.forEach(function(e){(map[e.address]=map[e.address]||[]).push(e)});
  var hit=0,seen={};
  function uniq(a){var s={};return a.filter(function(v){var k=String(v).toLowerCase().replace(/\s+/g," ").trim();if(!k||s[k])return false;s[k]=1;return true})}
  rows.forEach(function(x){
    var a=x.address.toLowerCase(),v=map[a];
    if(v){
      x.voter=uniq(v.map(function(e){return e.name})).join(", ");
      x.grp=uniq(v.map(function(e){return e.group})).join(", ");
      x.note=uniq(v.map(function(e){return e.note})).join(" | ");
      hit++;seen[a]=1;
    }else{x.voter="";x.grp="";x.note=""}
  });
  var all=Object.keys(map),miss=all.filter(function(a){return !seen[a]}).length;
  smsg("ตรงกัน "+hit.toLocaleString()+" จาก "+rows.length.toLocaleString()+" ผู้ถือ · address ในไฟล์ที่ไม่พบในผู้ถือ "+miss.toLocaleString()+" จาก "+all.length.toLocaleString()+" address");
  apply();
}
function doMatch(quiet){
  if(ge6)return doMatchGE6(quiet);
  var ca=$("scaddr").value,cn=$("scname").value,map={},sheetAddrs=0;
  sheetObjs.forEach(function(r){
    var m=String(r[ca]).match(ADDR);if(!m)return;
    var a=m[0].toLowerCase(),n=String(r[cn]).trim();sheetAddrs++;
    if(!map[a])map[a]=[];
    if(n&&map[a].indexOf(n)<0)map[a].push(n);
  });
  var hit=0,seen={};
  rows.forEach(function(x){
    var a=x.address.toLowerCase(),v=map[a];
    x.voter=v?v.join(", "):"";
    if(v){hit++;seen[a]=1}
  });
  var miss=Object.keys(map).filter(function(a){return !seen[a]}).length;
  if(!quiet||rows.length)smsg("ตรงกัน "+hit.toLocaleString()+" จาก "+rows.length.toLocaleString()+" ผู้ถือ · address ในชีตที่ไม่พบในผู้ถือ "+miss.toLocaleString()+" รายการ (จากทั้งหมด "+Object.keys(map).length.toLocaleString()+" address)");
  apply();
}
$("smatch").onclick=function(){
  if(!rows.length){smsg("กรุณากด โหลดข้อมูลทั้งหมดจาก API ก่อน",true);return}
  doMatch(false);
};

//GE6-PARSE-START
function parseGE6(sheets){
  var A=/0x[0-9a-fA-F]{40}/g,out=[];
  function txt(v){return String(v==null?"":v).replace(/\s+/g," ").trim()}
  sheets.forEach(function(sh){
    var a=sh.aoa,hdr=[];
    for(var r=0;r<a.length;r++)for(var c=0;c<a[r].length;c++){
      if(/^wallet\s*note$/i.test(txt(a[r][c])))hdr.push([r,c]);
    }
    hdr.forEach(function(h){
      var hr=h[0],wc=h[1],mc=-1,vc=-1;
      for(var c=wc-1;c>=0;c--){if(txt(a[hr][c])==="สมาชิก"){mc=c;break}}
      if(mc<0)return;
      for(var c2=wc+1;c2<a[hr].length&&c2<=wc+2;c2++){if(/จำนวนโหวต|โหวตไป/.test(txt(a[hr][c2]))){vc=c2;break}}
      var cur="";
      for(var r2=hr+1;r2<a.length;r2++){
        if(/^wallet\s*note$/i.test(txt(a[r2][wc])))break;
        var nm=txt(a[r2][mc]);
        if(nm){cur=(/^[-—–]+$/.test(nm))?"":nm}
        var cell=txt(a[r2][wc]),m=cell.match(A);
        if(!m||!cur)continue;
        var note=cell.replace(A,"").replace(/[()]/g," ").replace(/\s+/g," ").trim();
        m.forEach(function(ad){
          out.push({address:ad.toLowerCase(),name:cur,group:sh.name,note:note,votes:vc>=0?txt(a[r2][vc]):""});
        });
      }
    });
  });
  return out;
}
//GE6-PARSE-END
