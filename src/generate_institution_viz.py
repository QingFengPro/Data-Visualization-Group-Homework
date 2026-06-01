"""
生成合作机构排名对比图：中国机构 Top15 vs 中东欧机构 Top15。
数据源：使用完整明细文件（125/135 分别），而非截断的合并排名文件。
"""
import json, os
import pandas as pd

from common import (
    BASE_DIR, DIR_INSTITUTION, CEE_EXACT, TECH_BLUE_THEME as T,
    read_excel_safe, write_html, write_common_css, OUTPUT_DIR,
)

write_common_css()

# ── 读取中国机构完整明细 ──────────────────────────────────────
df_cn125 = read_excel_safe(
    os.path.join(DIR_INSTITUTION, "125中国机构.xlsx"),
    description="125中国机构明细",
)
df_cn135 = read_excel_safe(
    os.path.join(DIR_INSTITUTION, "135中国机构.xlsx"),
    description="135中国机构明细",
)

cn_all = {}
for _, row in df_cn125.iterrows():
    en = str(row.iloc[0]).strip()
    cn = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else en
    cnt = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
    cn_all[en] = {"name": cn, "name_en": en, "p125": cnt, "p135": 0}

for _, row in df_cn135.iterrows():
    en = str(row.iloc[0]).strip()
    cn = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else en
    cnt = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
    if en in cn_all:
        cn_all[en]["p135"] = cnt
        # Use the newer Chinese name if available
        if cn != en:
            cn_all[en]["name"] = cn
    else:
        cn_all[en] = {"name": cn, "name_en": en, "p125": 0, "p135": cnt}

# Compute rankings independently for each mode
cn_total = sorted([i for i in cn_all.values() if i["p125"] + i["p135"] > 0], key=lambda x: x["p125"] + x["p135"], reverse=True)
cn_125 = sorted([i for i in cn_all.values() if i["p125"] > 0], key=lambda x: x["p125"], reverse=True)
cn_135 = sorted([i for i in cn_all.values() if i["p135"] > 0], key=lambda x: x["p135"], reverse=True)

for i, inst in enumerate(cn_total):
    inst["rank_total"] = i + 1
for i, inst in enumerate(cn_125):
    inst["rank_125"] = i + 1
for i, inst in enumerate(cn_135):
    inst["rank_135"] = i + 1

# For institutions not in a particular ranking, assign a fallback rank
for inst in cn_all.values():
    if "rank_total" not in inst:
        inst["rank_total"] = len(cn_total) + 1
    if "rank_125" not in inst:
        inst["rank_125"] = len(cn_125) + 1
    if "rank_135" not in inst:
        inst["rank_135"] = len(cn_135) + 1
    inst["rank_change"] = inst["rank_125"] - inst["rank_135"]
    inst["delta"] = inst["p135"] - inst["p125"]

cn_top = cn_total[:15]

# ── 读取中东欧机构完整明细 ──────────────────────────────────
df_cee125 = read_excel_safe(
    os.path.join(DIR_INSTITUTION, "125中东欧机构.xlsx"),
    description="125中东欧机构明细",
)
df_cee135 = read_excel_safe(
    os.path.join(DIR_INSTITUTION, "135中东欧机构.xlsx"),
    description="135中东欧机构明细",
)

cee_all = {}
for _, row in df_cee125.iterrows():
    en = str(row.iloc[1]).strip()
    cn = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else en
    cnt = int(row.iloc[3]) if pd.notna(row.iloc[3]) else 0
    cty_en = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else ""
    cty_cn = CEE_EXACT.get(cty_en.upper(), cty_en)
    cee_all[en] = {"name": cn, "name_en": en, "country": cty_cn, "p125": cnt, "p135": 0}

for _, row in df_cee135.iterrows():
    en = str(row.iloc[1]).strip()
    cn = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else en
    cnt = int(row.iloc[3]) if pd.notna(row.iloc[3]) else 0
    cty_en = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else ""
    cty_cn = row.iloc[6] if len(row) > 6 and pd.notna(row.iloc[6]) else CEE_EXACT.get(cty_en.upper(), cty_en)
    if en in cee_all:
        cee_all[en]["p135"] = cnt
        if cn != en:
            cee_all[en]["name"] = cn
    else:
        cee_all[en] = {"name": cn, "name_en": en, "country": str(cty_cn), "p125": 0, "p135": cnt}

# Compute rankings independently for each mode
cee_total = sorted([i for i in cee_all.values() if i["p125"] + i["p135"] > 0], key=lambda x: x["p125"] + x["p135"], reverse=True)
cee_125 = sorted([i for i in cee_all.values() if i["p125"] > 0], key=lambda x: x["p125"], reverse=True)
cee_135 = sorted([i for i in cee_all.values() if i["p135"] > 0], key=lambda x: x["p135"], reverse=True)

for i, inst in enumerate(cee_total):
    inst["rank_total"] = i + 1
for i, inst in enumerate(cee_125):
    inst["rank_125"] = i + 1
for i, inst in enumerate(cee_135):
    inst["rank_135"] = i + 1

# For institutions not in a particular ranking, assign a fallback rank
for inst in cee_all.values():
    if "rank_total" not in inst:
        inst["rank_total"] = len(cee_total) + 1
    if "rank_125" not in inst:
        inst["rank_125"] = len(cee_125) + 1
    if "rank_135" not in inst:
        inst["rank_135"] = len(cee_135) + 1
    inst["rank_change"] = inst["rank_125"] - inst["rank_135"]
    inst["delta"] = inst["p135"] - inst["p125"]

cee_top = cee_total[:15]

print(f"CN institutions (total>0): {len(cn_total)}, 125>0: {len(cn_125)}, 135>0: {len(cn_135)}")
print(f"CEE institutions (total>0): {len(cee_total)}, 125>0: {len(cee_125)}, 135>0: {len(cee_135)}")
print(f"CN top 15: {[d['name'] for d in cn_top]}")
print(f"CEE top 15: {[d['name'] for d in cee_top]}")

html = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>合作机构排名</title>
<script src="https://d3js.org/d3.v7.min.js"></script>
<link rel="stylesheet" href="common.css">
<style>
.main{display:grid;grid-template-columns:1fr 1fr;gap:24px;max-width:1400px;width:95%;margin:20px auto}
.panel{background:#fff;border:1px solid var(--c-border-lt);border-radius:16px;padding:24px 20px 18px;box-shadow:0 2px 12px rgba(0,0,0,0.06);min-height:480px}
.panel-head{display:flex;align-items:baseline;gap:8px;margin-bottom:4px}
.panel h3{font-size:16px;font-weight:700;margin:0}
.panel .sub{font-size:11px;color:#999;margin-bottom:16px}
.chart{width:100%}
.c-blue{color:var(--c-primary)}.c-purple{color:#3949ab}
.rank-change{display:inline-block;font-size:10px;margin-left:4px}
.rank-change.up{color:#2e7d32}
.rank-change.down{color:#c62828}
.rank-change.same{color:#999}
.chart-row{display:flex;align-items:center;gap:8px;padding:2px 0}
.chart-rank{flex:0 0 20px;text-align:center}
.chart-name{flex:0 0 150px;font-size:11px;color:#333;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.chart-bar-wrap{flex:1;height:22px;position:relative;display:flex;align-items:stretch}
.chart-bar-bg{position:absolute;inset:0;background:#f5f5f5;border-radius:4px}
.chart-bar-fill{position:absolute;top:0;height:100%;border-radius:4px;opacity:0.85;transition:width 0.7s}
.chart-bar-fill.positive{left:50%;background:#69f0ae}
.chart-bar-fill.negative{right:50%;background:#ff6e40}
.chart-bar-center{position:absolute;left:50%;top:0;width:1px;height:100%;background:#e0e0e0;transform:translateX(-0.5px)}
.chart-val{flex:0 0 65px;font-size:11px;color:#555;text-align:right;font-weight:500}
.chart-val.positive{color:#2e7d32}
.chart-val.negative{color:#c62828}
.tooltip-inner{padding:10px 14px}
.tooltip-inner b{color:var(--c-blue);font-size:13px}
.tooltip-inner .detail{font-size:11px;color:#ddd;line-height:1.7}
//.stats-summary{display:flex;justify-content:center;gap:32px;padding:14px 20px;background:#fff;border-bottom:1px solid var(--c-border-ft)}
//.stats-summary .item{text-align:center}
//.stats-summary .num{font-size:20px;font-weight:800;color:var(--c-primary)}
//.stats-summary .label{font-size:11px;color:#999;margin-top:2px}
@media(max-width:960px){.main{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="header"><h1>合作机构排名对比</h1><p>基于完整明细数据 | 中国 Top 15 vs 中东欧 Top 15 | 含 125/135 分期排名</p></div>
<div class="stats-summary" id="statsSummary"></div>
<div class="controls">
  <button class="mode-btn active" data-mode="total">总体 (125+135)</button>
  <div class="sep"></div>
  <button class="mode-btn" data-mode="125">125 期间</button>
  <button class="mode-btn" data-mode="135">135 期间</button>
  <button class="mode-btn delta" data-mode="delta">变化 (125→135)</button>
</div>
<div class="main">
  <div class="panel" id="panelCN">
    <div class="panel-head"><h3 class="c-blue">中国机构 Top 15</h3></div>
    <div class="sub" id="subCN">按 125+135 总发文量排序</div>
    <div class="chart" id="chartCN"></div>
  </div>
  <div class="panel" id="panelCEE">
    <div class="panel-head"><h3 class="c-purple">中东欧机构 Top 15</h3></div>
    <div class="sub" id="subCEE">按 125+135 总发文量排序 · 含国别</div>
    <div class="chart" id="chartCEE"></div>
  </div>
</div>
<div class="tooltip" id="tooltip"></div>
<script>
const CN = __CN_DATA__;
const CEE = __CEE_DATA__;
let currentMode = "total";

function getVal(d){
  if(currentMode==="125") return d.p125;
  if(currentMode==="135") return d.p135;
  if(currentMode==="delta") return d.delta;
  return d.p125+d.p135;
}
function getRank(d){
  if(currentMode==="125") return d.rank_125;
  if(currentMode==="135") return d.rank_135;
  return d.rank_total;
}
function getRankLabel(d){
  var rc=d.rank_change;
  if(rc>0) return {cls:"up",txt:"↑"+rc};
  if(rc<0) return {cls:"down",txt:"↓"+Math.abs(rc)};
  return {cls:"same",txt:"−"};
}
function filterData(data){
  if(currentMode==="125") return data.filter(d=>d.p125>0);
  if(currentMode==="135") return data.filter(d=>d.p135>0);
  if(currentMode==="delta") return data.filter(d=>d.p125>0&&d.p135>0);
  return data.filter(d=>d.p125+d.p135>0);
}

// d3.select("#statsSummary").html(
//   '<div class="item"><div class="num">'+CN.length+'</div><div class="label">中国机构(全量)</div></div>'+
//   '<div class="item"><div class="num">'+CEE.length+'</div><div class="label">中东欧机构(全量)</div></div>'+
//   '<div class="item"><div class="num">'+d3.sum(CN,d=>d.p125+d.p135).toLocaleString()+'</div><div class="label">中国总论文</div></div>'+
//   '<div class="item"><div class="num">'+d3.sum(CEE,d=>d.p125+d.p135).toLocaleString()+'</div><div class="label">中东欧总论文</div></div>'
// );

function drawPanel(containerId, data, accentColor, showCountry){
  var c=d3.select("#"+containerId);
  var isDelta=currentMode==="delta";
  var filtered=filterData(data);
  var sorted=[...filtered].sort((a,b)=>getVal(b)-getVal(a));
  var vals=sorted.map(d=>getVal(d));
  var maxV=isDelta?d3.max(vals.map(Math.abs))||1:d3.max(vals)||1;

  c.selectAll(".chart-row").data(sorted).join(
    enter=>{
      var row=enter.append("div").attr("class","chart-row").style("opacity",0);
      row.transition().duration(400).delay((d,i)=>i*40).style("opacity",1);

      var rk=row.append("span").attr("class","chart-rank");
      if(isDelta){
        rk.append("span").attr("class",function(){return ""}).style("display","none");
      }else{
        rk.append("span").attr("class",d=>"rank-badge"+(getRank(d)<=3?" top"+getRank(d) :"")).text((d,i)=>i+1);
      }

      var nm=row.append("span").attr("class","chart-name")
        .text(d=>{var n=d.name.length>12?d.name.slice(0,11)+"..":d.name;if(showCountry&&d.country)n+=" ("+d.country+")";return n;});
      if(isDelta){
        var rcSpan=nm.append("span").attr("class",d=>"rank-change "+getRankLabel(d).cls).style("margin-left","3px");
        rcSpan.text(d=>getRankLabel(d).txt);
      }

      var bw=row.append("span").attr("class","chart-bar-wrap");
      bw.append("span").attr("class","chart-bar-bg");
      if(isDelta){
        bw.append("span").attr("class","chart-bar-center");
        var barPct=d=>(Math.abs(getVal(d))/maxV*50)+"%";
        var positiveFill=bw.filter(d=>getVal(d)>=0).append("span").attr("class","chart-bar-fill positive").style("width","0%");
        var negativeFill=bw.filter(d=>getVal(d)<0).append("span").attr("class","chart-bar-fill negative").style("width","0%");
        setTimeout(function(){
          positiveFill.style("width",barPct);
          negativeFill.style("width",barPct);
        },50);
      }else{
        bw.append("span").attr("class","chart-bar-fill").style("background",accentColor).style("left","0").style("width",d=>(getVal(d)/maxV*100)+"%");
      }

      var valText=isDelta?(d=>(getVal(d)>=0?"+"+getVal(d):getVal(d)).toLocaleString()+" 篇"):(d=>getVal(d).toLocaleString()+" 篇");
      row.append("span").attr("class",d=>"chart-val"+(isDelta?(getVal(d)>=0?" positive":" negative"):"")).text(valText);

      row.on("mouseenter",function(e,d){
        d3.select(this).select(".chart-bar-fill").style("opacity",1);
        var html='<div class="tooltip-inner"><b>'+d.name+'</b>'+(d.name_en?'<br><span style="color:#aaa;font-size:10px">'+d.name_en+'</span>':'');
        if(showCountry&&d.country) html+='<br><span style="color:#aaa;font-size:10px">'+d.country+'</span>';
        html+='<div class="detail">'+
          '125: '+d.p125.toLocaleString()+' 篇 (第'+d.rank_125+')<br>'+
          '135: '+d.p135.toLocaleString()+' 篇 (第'+d.rank_135+')<br>'+
          '合计: '+(d.p125+d.p135).toLocaleString()+' 篇<br>'+
          '变化: '+(d.delta>=0?'+':'')+d.delta+' 篇<br>'+
          '排名: '+getRankLabel(d).txt;
        d3.select("#tooltip").html(html).style("opacity",1);
      })
      .on("mousemove",function(e){d3.select("#tooltip").style("left",(e.pageX+14)+"px").style("top",(e.pageY-10)+"px")})
      .on("mouseleave",function(){
        d3.select(this).select(".chart-bar-fill").style("opacity",0.85);
        d3.select("#tooltip").style("opacity",0);
      });
    }
  );
}

function renderAll(){
  var modeLabels={"total":"125+135","125":"125","135":"135","delta":"变化量 (125→135)"};
  var isDelta=currentMode==="delta";
  var cnFiltered=filterData(CN);
  var ceeFiltered=filterData(CEE);
  d3.select("#subCN").text((isDelta?"按变化量排序 | ":"按 "+modeLabels[currentMode]+" 排序 | ")+"Top 15 · 共 "+cnFiltered.length+" 个");
  d3.select("#subCEE").text((isDelta?"按变化量排序 | ":"按 "+modeLabels[currentMode]+" 排序 | ")+"含国别 | Top 15 · 共 "+ceeFiltered.length+" 个");
  var cnColor=isDelta?"#69f0ae":"#4fc3f7";
  var ceeColor=isDelta?"#69f0ae":"#7c4dff";
  drawPanel("chartCN", CN, cnColor, false);
  drawPanel("chartCEE", CEE, ceeColor, true);
}

d3.selectAll(".mode-btn").on("click",function(){
  currentMode=this.dataset.mode;
  d3.selectAll(".mode-btn").classed("active",function(){return this.dataset.mode===currentMode});
  d3.selectAll(".mode-btn.delta").classed("active",currentMode==="delta");
  d3.select("#chartCN").selectAll("*").remove();
  d3.select("#chartCEE").selectAll("*").remove();
  renderAll();
});

renderAll();
</script>
</body>
</html>"""

html = html.replace("__CN_DATA__", json.dumps(cn_top, ensure_ascii=False))
html = html.replace("__CEE_DATA__", json.dumps(cee_top, ensure_ascii=False))
write_html(os.path.join(OUTPUT_DIR, "institution_visualization.html"), html, "机构排名")
print(f"  中国机构 Top 15: {len(cn_top)}, 中东欧机构 Top 15: {len(cee_top)}")
print("Done.")
