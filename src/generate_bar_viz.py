"""
读取我国与各国在125-135所发合作文章数量.xlsx，生成横向条形图 + 国家简介。
科技深蓝主题 + 入场动画 + 变化量视图。
"""
import json, os
import pandas as pd

from common import (
    BASE_DIR, DIR_COUNTRY, CEE_EXACT, TECH_BLUE_THEME as T,
    read_excel_safe, write_html, write_common_css, OUTPUT_DIR,
)

write_common_css()

xlsx_path = os.path.join(DIR_COUNTRY, "我国与各国在125-135所发合作文章数量.xlsx")
df = read_excel_safe(xlsx_path, description="各国发文量")

# Build lookup tables from both 125 and 135 columns
p125_map = {}
p135_map = {}
for _, row in df.iterrows():
    name125 = str(row.iloc[0]).strip().upper()
    if name125 in CEE_EXACT:
        p125_map[name125] = {
            "p125": int(row.iloc[1] or 0),
            "r125": int(row.iloc[2]) if pd.notna(row.iloc[2]) else "-",
        }
    name135 = str(row.iloc[5]).strip().upper()
    if name135 in CEE_EXACT:
        p135_map[name135] = {
            "p135": int(row.iloc[6] or 0),
            "r135": int(row.iloc[7]) if pd.notna(row.iloc[7]) else "-",
        }

cee_data = []
for name in CEE_EXACT:
    d125 = p125_map.get(name, {"p125": 0, "r125": "-"})
    d135 = p135_map.get(name, {"p135": 0, "r135": "-"})
    p125 = d125["p125"]
    p135 = d135["p135"]
    cee_data.append({
        "name_cn": CEE_EXACT[name], "name_en": name,
        "p125": p125, "r125": d125["r125"],
        "p135": p135, "r135": d135["r135"],
        "delta": p135 - p125,
    })

# Sort by 135 period by default
cee_data.sort(key=lambda x: x["p135"], reverse=True)
print(f"Loaded {len(cee_data)} CEE countries")

PROFILES = {
    "波兰": {"capital": "华沙","area": "31.27万 km²","population": "约 3,775 万 (2023)","location": "中欧，北濒波罗的海，东邻白俄罗斯、乌克兰，南接捷克、斯洛伐克，西连德国","language": "波兰语","currency": "兹罗提 (PLN)","eu": "欧盟成员国 (2004年加入)"},
    "捷克": {"capital": "布拉格","area": "7.89万 km²","population": "约 1,087 万 (2023)","location": "中欧内陆国，北邻波兰，东接斯洛伐克，南连奥地利，西毗德国","language": "捷克语","currency": "捷克克朗 (CZK)","eu": "欧盟成员国 (2004年加入)"},
    "希腊": {"capital": "雅典","area": "13.20万 km²","population": "约 1,036 万 (2023)","location": "东南欧，巴尔干半岛南端，濒爱琴海、爱奥尼亚海和地中海","language": "希腊语","currency": "欧元 (EUR)","eu": "欧盟成员国 (1981年加入)"},
    "匈牙利": {"capital": "布达佩斯","area": "9.30万 km²","population": "约 960 万 (2023)","location": "中欧内陆国，北邻斯洛伐克，东接乌克兰、罗马尼亚，南连塞尔维亚、克罗地亚、斯洛文尼亚","language": "匈牙利语","currency": "福林 (HUF)","eu": "欧盟成员国 (2004年加入)"},
    "罗马尼亚": {"capital": "布加勒斯特","area": "23.84万 km²","population": "约 1,905 万 (2023)","location": "东南欧，巴尔干半岛北部，濒黑海","language": "罗马尼亚语","currency": "列伊 (RON)","eu": "欧盟成员国 (2007年加入)"},
    "塞尔维亚": {"capital": "贝尔格莱德","area": "8.84万 km²","population": "约 667 万 (2023)","location": "东南欧，巴尔干半岛中部内陆国","language": "塞尔维亚语","currency": "第纳尔 (RSD)","eu": "欧盟候选国"},
    "斯洛文尼亚": {"capital": "卢布尔雅那","area": "2.03万 km²","population": "约 212 万 (2023)","location": "中欧南部，北邻奥地利，东接匈牙利，南连克罗地亚，西毗意大利","language": "斯洛文尼亚语","currency": "欧元 (EUR)","eu": "欧盟成员国 (2004年加入)"},
    "克罗地亚": {"capital": "萨格勒布","area": "5.66万 km²","population": "约 385 万 (2023)","location": "东南欧，巴尔干半岛西北部，濒亚得里亚海","language": "克罗地亚语","currency": "欧元 (EUR)","eu": "欧盟成员国 (2013年加入)"},
    "斯洛伐克": {"capital": "布拉迪斯拉发","area": "4.90万 km²","population": "约 543 万 (2023)","location": "中欧内陆国，北邻波兰，东接乌克兰，南连匈牙利，西毗奥地利、捷克","language": "斯洛伐克语","currency": "欧元 (EUR)","eu": "欧盟成员国 (2004年加入)"},
    "爱沙尼亚": {"capital": "塔林","area": "4.52万 km²","population": "约 136 万 (2023)","location": "东北欧，波罗的海东岸，北濒芬兰湾","language": "爱沙尼亚语","currency": "欧元 (EUR)","eu": "欧盟成员国 (2004年加入)"},
    "保加利亚": {"capital": "索非亚","area": "11.10万 km²","population": "约 644 万 (2023)","location": "东南欧，巴尔干半岛东部，濒黑海","language": "保加利亚语","currency": "列弗 (BGN)","eu": "欧盟成员国 (2007年加入)"},
    "立陶宛": {"capital": "维尔纽斯","area": "6.53万 km²","population": "约 287 万 (2023)","location": "东北欧，波罗的海东岸","language": "立陶宛语","currency": "欧元 (EUR)","eu": "欧盟成员国 (2004年加入)"},
    "拉脱维亚": {"capital": "里加","area": "6.46万 km²","population": "约 188 万 (2023)","location": "东北欧，波罗的海东岸","language": "拉脱维亚语","currency": "欧元 (EUR)","eu": "欧盟成员国 (2004年加入)"},
    "北马其顿": {"capital": "斯科普里","area": "2.57万 km²","population": "约 206 万 (2023)","location": "东南欧，巴尔干半岛中部内陆国","language": "马其顿语","currency": "代纳尔 (MKD)","eu": "欧盟候选国"},
    "波黑": {"capital": "萨拉热窝","area": "5.12万 km²","population": "约 322 万 (2023)","location": "东南欧，巴尔干半岛中西部","language": "波斯尼亚语/克罗地亚语/塞尔维亚语","currency": "可兑换马克 (BAM)","eu": "欧盟候选国"},
    "黑山": {"capital": "波德戈里察","area": "1.38万 km²","population": "约 62 万 (2023)","location": "东南欧，巴尔干半岛西部，濒亚得里亚海","language": "黑山语","currency": "欧元 (EUR)","eu": "欧盟候选国"},
    "阿尔巴尼亚": {"capital": "地拉那","area": "2.87万 km²","population": "约 277 万 (2023)","location": "东南欧，巴尔干半岛西部，濒亚得里亚海和爱奥尼亚海","language": "阿尔巴尼亚语","currency": "列克 (ALL)","eu": "欧盟候选国"},
}

html = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>各国合作论文量与排名</title>
<script src="https://d3js.org/d3.v7.min.js"></script>
<link rel="stylesheet" href="common.css">
<style>
.main-container{display:flex;gap:20px;padding:20px;max-width:1300px;margin:0 auto;min-height:600px}
.chart-panel{flex:1;background:#fff;border:1px solid #e8e8e8;border-radius:14px;padding:20px;box-shadow:0 2px 12px rgba(0,0,0,0.06)}
.chart-title{font-size:14px;font-weight:600;color:var(--c-primary);margin-bottom:12px}
.profile-panel{flex:0 0 340px;background:#fff;border:1px solid #e8e8e8;border-radius:14px;padding:24px;box-shadow:0 2px 12px rgba(0,0,0,0.06)}
.profile-title{font-size:17px;font-weight:700;color:var(--c-primary);margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid #d0d0d0}
.profile-placeholder{color:#bbb;font-size:14px;text-align:center;padding-top:100px}
.profile-placeholder span{display:block;font-size:44px;margin-bottom:10px}
.profile-field{margin-bottom:12px}
.profile-label{font-size:11px;color:#bbb;margin-bottom:2px}
.profile-value{font-size:14px;color:#666;line-height:1.5}
</style>
</head>
<body>
<div class="header"><h1>中东欧各国与我国合作发文量</h1><p>数据来源: Web of Science 125–135 | 点击条形查看国家简介</p></div>
<div class="controls">
  <button class="mode-btn active" data-mode="125">125 期间</button>
  <button class="mode-btn" data-mode="135">135 期间</button>
  <button class="mode-btn delta" data-mode="delta">变化 (125→135)</button>
</div>
<div class="main-container">
<div class="chart-panel"><div class="chart-title" id="chartTitle">合作发文量及全球排名 — 125 期间</div><div id="barChart"></div></div>
<div class="profile-panel" id="profilePanel"><div class="profile-placeholder"><span>&#128269;</span><div>点击左侧条形图查看该国简介</div></div></div>
</div>
<div class="tooltip" id="tooltip"></div>
<script>
const CEE_DATA=__CEE_DATA__, PROFILES=__PROFILES__;
let currentMode="125", selectedCountry=null;

function getVal(d){return currentMode==="125"?d.p125:currentMode==="135"?d.p135:d.delta}
function getRank(d){return currentMode==="delta"?"":" (全球第 "+(currentMode==="125"?d.r125:d.r135)+" 位)"}

function drawChart(){
  d3.select("#barChart").selectAll("*").remove();
  const data=[...CEE_DATA].sort((a,b)=>getVal(b)-getVal(a));
  const vals=data.map(d=>getVal(d));
  const absMax=d3.max(vals.map(Math.abs))||1;
  const w=580,barH=28,gap=7,top=8,bot=20;
  const h=data.length*(barH+gap)+top+bot;
  const svg=d3.select("#barChart").append("svg").attr("width",w).attr("height",h);
  const cx=130;
  const badgeR=10;
  const badgeColors=["#ffd740","#b0bec5","#bcaaa4"];
  let xS;
  if(currentMode==="delta") xS=d3.scaleLinear().domain([-absMax*1.15,absMax*1.15]).range([0,w-cx-80]);
  else xS=d3.scaleLinear().domain([0,(d3.max(vals)||1)*1.12]).range([0,w-cx-80]);

  const rows=svg.selectAll("g").data(data).join("g").attr("transform",(d,i)=>`translate(${cx},${top+i*(barH+gap)})`);

  // bg
  rows.append("rect").attr("x",0).attr("y",0).attr("width",w-cx-80).attr("height",barH).attr("rx",4).attr("fill","#fafafa");

  // delta zero line
  if(currentMode==="delta") svg.append("line").attr("x1",cx+xS(0)).attr("x2",cx+xS(0)).attr("y1",top).attr("y2",top+data.length*(barH+gap)).attr("stroke","#e0e0e0").attr("stroke-width",1).attr("stroke-dasharray","5,3");

  // value bar (with enter animation)
  rows.append("rect").attr("class","val-bar")
    .attr("x",d=>currentMode==="delta"?(getVal(d)>=0?xS(0):xS(getVal(d))):0)
    .attr("y",0).attr("width",0).attr("height",barH).attr("rx",4)
    .attr("fill",d=>currentMode==="delta"?(getVal(d)>=0?"#69f0ae":"#ff6e40"):"#4fc3f7")
    .on("mouseenter",function(e,d){d3.select(this).attr("opacity",1);const dl=d.delta>=0?"+"+d.delta:d.delta;showTooltip(e,`<b>${d.name_cn}</b><br>${getValLabel(d)}${currentMode!=="delta"?"<br>变化: "+dl+" 篇":""}`)})
    .on("mouseleave",function(e,d){d3.select(this).attr("opacity",0.85);hideTooltip()})
    .on("click",(e,d)=>{selectedCountry=selectedCountry&&selectedCountry.name_cn===d.name_cn?null:d;d3.selectAll("rect.val-bar").attr("fill",dd=>{if(currentMode==="delta")return getVal(dd)>=0?"#69f0ae":"#ff6e40";return selectedCountry&&dd.name_cn===selectedCountry.name_cn?"#00e5ff":"#4fc3f7"});showProfile(selectedCountry)})
    .transition().duration(700).delay((d,i)=>i*45)
    .attr("width",d=>currentMode==="delta"?Math.abs(xS(getVal(d))-xS(0)):Math.max(xS(getVal(d)),3)).attr("opacity",0.85);

  // rank badges + names
  const nameG=svg.selectAll("g.nm").data(data).join("g").attr("class","nm")
    .attr("transform",(d,i)=>`translate(${cx-12},${top+i*(barH+gap)+barH/2})`);
  nameG.append("circle").attr("r",badgeR).attr("fill",(d,i)=>i<3?badgeColors[i]:"#e0e0e0").attr("stroke",(d,i)=>i<3?badgeColors[i]:"none").attr("stroke-width",1);
  nameG.append("text").attr("text-anchor","middle").attr("dy","0.38em").style("fill",(d,i)=>i<3?"#0a0e27":"#fff").style("font-size","10px").style("font-weight","700").text((d,i)=>i+1);
  nameG.append("text").attr("x",-badgeR-6).attr("text-anchor","end").attr("dy","0.38em").style("fill","#333").style("font-size","13px").style("font-weight","700").text(d=>d.name_cn);

  // values
  svg.selectAll("text.vl").data(data).join("text").attr("class","vl")
    .attr("x",d=>{const v=getVal(d);return cx+(currentMode==="delta"?(xS(v)+(v>=0?24:-24)):xS(v)+6)})
    .attr("y",(d,i)=>top+i*(barH+gap)+barH/2+4)
    .attr("text-anchor",currentMode==="delta"?"middle":"start")
    .style("fill",d=>currentMode==="delta"?(getVal(d)>=0?"#2e7d32":"#c62828"):"#555").style("font-size","11px")
    .text(d=>getValLabel(d));

  d3.select("#chartTitle").text(`合作发文量及全球排名 — ${currentMode==="125"?"125 期间":currentMode==="135"?"135 期间":"变化量 (125→135)"}`);
}

function getValLabel(d){const v=getVal(d),r=getRank(d);return currentMode==="delta"?(v>=0?"+"+v:v)+" 篇":v+" 篇"+r}

function showProfile(d){
  const p=d3.select("#profilePanel");p.selectAll("*").remove();
  if(!d){p.append("div").attr("class","profile-placeholder").html(`<span>&#128269;</span><div>点击左侧条形图查看该国简介</div>`);return}
  const pr=PROFILES[d.name_cn];if(!pr)return;
  p.append("div").attr("class","profile-title").text(" "+d.name_cn);
  [["首都",pr.capital],["地理位置",pr.location],["国土面积",pr.area],["人口",pr.population],["官方语言",pr.language],["货币",pr.currency],["欧盟状态",pr.eu],["125 期间",`${d.p125} 篇 (全球第 ${d.r125} 位)`],["135 期间",`${d.p135} 篇 (全球第 ${d.r135} 位)`],["变化",`${d.delta>=0?'+':''}${d.delta} 篇`]].forEach(([l,v])=>{const r=p.append("div").attr("class","profile-field");r.append("div").attr("class","profile-label").text(l);r.append("div").attr("class","profile-value").text(v)});
}

function showTooltip(e,h){d3.select("#tooltip").html(h).style("opacity",1).style("left",(e.pageX+14)+"px").style("top",(e.pageY-10)+"px")}
function hideTooltip(){d3.select("#tooltip").style("opacity",0)}

d3.selectAll(".mode-btn").on("click",function(){currentMode=this.dataset.mode;d3.selectAll(".mode-btn").classed("active",function(){return this.dataset.mode===currentMode});d3.selectAll(".mode-btn.delta").classed("active",currentMode==="delta");drawChart()});

drawChart();
</script>
</body>
</html>"""

html = html.replace("__CEE_DATA__", json.dumps(cee_data, ensure_ascii=False))
html = html.replace("__PROFILES__", json.dumps(PROFILES, ensure_ascii=False))
write_html(os.path.join(OUTPUT_DIR, "cee_bar_ranking.html"), html, "国家排名条形图")
print(f"  Countries: {len(cee_data)}")
print("Done.")
