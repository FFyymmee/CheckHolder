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
  var btn=$("go");btn.disabled=true;rows=[];var params="",n=0,guard=0;
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
        var bal=Number(it.value)/Math.pow(10,d);
        rows.push({address:it.address.hash,type:it.address.is_contract?"Contract":"Wallet",balance:bal,pct:info.supply>0?bal/info.supply*100:NaN});
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
    sk="balance";sd=-1;page=0;$("out").hidden=false;stats();apply();
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
  var q=$("q").value.trim().toLowerCase();
  view=rows.filter(function(x){return !q||x.address.toLowerCase().indexOf(q)>-1});
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
    return '<tr><td class="n">'+x.rank+'</td><td class="a">'+x.address+'</td><td>'+x.type+'</td><td class="n">'+fmt(x.balance)+'</td><td class="n">'+(isNaN(x.pct)?"-":x.pct.toFixed(4)+"%")+'</td></tr>';
  }).join("");
  $("pi").textContent="หน้า "+(page+1)+" / "+pages+" ("+view.length.toLocaleString()+")";
}
document.querySelectorAll("th").forEach(function(th){th.onclick=function(){
  var k=th.dataset.k;if(sk===k)sd=-sd;else{sk=k;sd=(k==="address"||k==="type")?1:-1}apply()}});
$("q").oninput=function(){page=0;apply()};
$("prev").onclick=function(){if(page>0){page--;render()}};
$("next").onclick=function(){page++;render()};
function data(){return rows.map(function(x){return {"Rank":x.rank,"Address":x.address,"Type":x.type,"Balance":x.balance,"Percent of supply":isNaN(x.pct)?"":Number(x.pct.toFixed(6))}})}
$("xl").onclick=function(){
  var wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data()),"Holders");
  XLSX.writeFile(wb,(info.symbol||"token")+"-holders.xlsx");
};
$("csv").onclick=function(){
  var csv=XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(data()));
  var a=document.createElement("a");a.href=URL.createObjectURL(new Blob(["\ufeff"+csv],{type:"text/csv"}));
  a.download=(info.symbol||"token")+"-holders.csv";a.click();
};
