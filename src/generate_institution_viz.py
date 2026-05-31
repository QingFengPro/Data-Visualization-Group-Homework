"""
生成合作机构排名对比图：中国机构 Top15 vs 中东欧机构 Top15。
横向条形图，带排名变化指示。
"""
import json, os
import pandas as pd

from common import (
    BASE_DIR, DIR_INSTITUTION, CN_TO_ISO, TECH_BLUE_THEME as T,
    read_excel_safe, write_html, write_common_css, OUTPUT_DIR,
)

write_common_css()

# ── 读取中国机构 ──────────────────────────────────────────────
cn_path = os.path.join(DIR_INSTITUTION, "125-135中国机构所发合作文章数量.xlsx")
df_cn = read_excel_safe(cn_path, description="中国机构")
cn_insts = []
for _, row in df_cn.iterrows():
    try:
        cn_insts.append({
            "rank_125": int(row.iloc[6]) if pd.notna(row.iloc[6]) else None,
            "name": str(row.iloc[1]).strip(),
            "p135": int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0,
            "p125": int(row.iloc[4]) if pd.notna(row.iloc[4]) else 0,
            "delta": int(row.iloc[7]) if pd.notna(row.iloc[7]) else 0,
        })
    except (ValueError, IndexError):
        continue

cn_insts.sort(key=lambda x: x["p135"] + x["p125"], reverse=True)
cn_top = cn_insts[:15]

# Compute ranks within dataset
cn_by_p125 = sorted(cn_insts, key=lambda x: x["p125"], reverse=True)
cn_by_p135 = sorted(cn_insts, key=lambda x: x["p135"], reverse=True)
cn_r135 = {inst["name"]: i+1 for i, inst in enumerate(cn_by_p135)}
cn_r125 = {inst["name"]: i+1 for i, inst in enumerate(cn_by_p125)}
for inst in cn_insts:
    inst["rank_135"] = cn_r135[inst["name"]]
    inst["rank_change"] = cn_r125[inst["name"]] - cn_r135[inst["name"]]

# ── 读取中东欧机构 ────────────────────────────────────────────
cee_path = os.path.join(DIR_INSTITUTION, "125-135中东欧机构所发合作文章数量.xlsx")
df_cee = read_excel_safe(cee_path, description="中东欧机构")
cee_insts = []
for _, row in df_cee.iterrows():
    try:
        cee_insts.append({
            "name": str(row.iloc[1]).strip(),
            "country": str(row.iloc[2]).strip(),
            "p135": int(row.iloc[3]) if pd.notna(row.iloc[3]) else 0,
            "p125": int(row.iloc[5]) if pd.notna(row.iloc[5]) else 0,
            "rank_125": int(row.iloc[7]) if pd.notna(row.iloc[7]) else None,
            "delta": int(row.iloc[9]) if pd.notna(row.iloc[9]) else 0,
        })
    except (ValueError, IndexError):
        continue

cee_insts.sort(key=lambda x: x["p135"] + x["p125"], reverse=True)
cee_top = cee_insts[:15]

# Compute ranks within dataset
cee_by_p125 = sorted(cee_insts, key=lambda x: x["p125"], reverse=True)
cee_by_p135 = sorted(cee_insts, key=lambda x: x["p135"], reverse=True)
cee_r135 = {inst["name"]: i+1 for i, inst in enumerate(cee_by_p135)}
cee_r125 = {inst["name"]: i+1 for i, inst in enumerate(cee_by_p125)}
for inst in cee_insts:
    inst["rank_135"] = cee_r135[inst["name"]]
    inst["rank_change"] = cee_r125[inst["name"]] - cee_r135[inst["name"]]

print(f"Chinese institutions: {len(cn_top)}, CEE institutions: {len(cee_top)}")

html = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>合作机构排名</title>
<script src="https://d3js.org/d3.v7.min.js"></script>
<link rel="stylesheet" href="common.css">
<style>
.main{display:flex;gap:28px;max-width:1300px;width:95%;margin:20px auto}
.panel{flex:1;background:#fff;border:1px solid #e0e0e0;border-radius:16px;padding:28px 22px 20px;box-shadow:0 2px 12px rgba(0,0,0,0.08)}
.panel h3{font-size:16px;font-weight:600;margin-bottom:4px}
.panel .sub{font-size:11px;color:#999;margin-bottom:14px}
.chart{width:100%}
.c-blue{color:var(--c-primary)}.c-purple{color:#3949ab}
</style>
</head>
<body>
<div class="header"><h1>合作机构排名对比</h1><p>中国机构 Top 15 vs 中东欧机构 Top 15 | 点击切换时期</p></div>
<div class="controls">
  <button class="mode-btn active" data-mode="total">总体 (125+135)</button>
  <button class="mode-btn" data-mode="125">125 期间</button>
  <button class="mode-btn" data-mode="135">135 期间</button>
</div>
<div class="main">
<div class="panel"><h3 class="c-blue">🏛️ 中国机构 Top 15</h3><div class="sub" id="subCN">按 125+135 总发文量排序</div><div class="chart" id="chartCN"></div></div>
<div class="panel"><h3 class="c-purple">🏛️ 中东欧机构 Top 15</h3><div class="sub" id="subCEE">按 125+135 总发文量排序 | 含国别</div><div class="chart" id="chartCEE"></div></div>
</div>
<div class="tooltip" id="tooltip"></div>
<script>
const CN = __CN_DATA__;
const CEE = __CEE_DATA__;
let currentMode = "total";

function getVal(d) {
  if(currentMode==="125") return d.p125;
  if(currentMode==="135") return d.p135;
  return d.p125 + d.p135;
}

function drawBar(containerId, data, accentColor, showCountry) {
  const c = d3.select("#"+containerId);
  const w = c.node().clientWidth;
  const barH=26, gap=6, top=8, bot=16;
  const sorted = [...data].sort((a,b)=>getVal(b)-getVal(a));
  const h = sorted.length*(barH+gap)+top+bot;
  const svg = c.append("svg").attr("width",w).attr("height",h);
  const maxV = d3.max(sorted, d=>getVal(d)) || 1;
  const x = d3.scaleLinear().domain([0, maxV*1.12]).range([0, w-220]);

  const rows = svg.selectAll("g").data(sorted).join("g")
    .attr("transform",(d,i)=>`translate(0,${top+i*(barH+gap)})`);

  // rank number
  rows.append("text").attr("x",0).attr("y",barH/2+4).attr("text-anchor","start")
    .style("fill","#8899aa").style("font-size","12px").style("font-weight","700")
    .text((d,i)=>i+1);

  // name (with country for CEE)
  rows.append("text").attr("x",28).attr("y",barH/2+4).attr("text-anchor","start")
    .style("fill","#333").style("font-size","11px")
    .text(d => {
      let n = d.name.length>18 ? d.name.slice(0,16)+".." : d.name;
      if(showCountry && d.country) n += "  "+d.country;
      return n;
    });

  // bg bar
  rows.append("rect").attr("x",195).attr("y",0).attr("width",w-220).attr("height",barH)
    .attr("rx",3).attr("fill","#f5f5f5");

  // value bar
  rows.append("rect").attr("x",195).attr("y",0).attr("width",0).attr("height",barH).attr("rx",3)
    .attr("fill",accentColor).attr("opacity",0.8)
    .transition().duration(800).delay((d,i)=>i*50)
    .attr("width",d=>Math.max(x(getVal(d)),3));

  // value text
  rows.append("text").attr("x",d=>200+Math.max(x(getVal(d)),3)+6).attr("y",barH/2+5)
    .style("fill","#555").style("font-size","10px")
    .text(d=>{
      if(currentMode==="125") return d.p125+" 篇";
      if(currentMode==="135") return d.p135+" 篇";
      return (d.p125+d.p135)+" 篇";
    });


}

function renderAll() {
  d3.select("#chartCN").selectAll("*").remove();
  d3.select("#chartCEE").selectAll("*").remove();
  const modeLabels = {"total":"125+135","125":"125","135":"135"};
  d3.select("#subCN").text("按 "+modeLabels[currentMode]+" 排序 | Top 15");
  d3.select("#subCEE").text("按 "+modeLabels[currentMode]+" 排序 | 含国别 | Top 15");
  drawBar("chartCN", CN, "#4fc3f7", false);
  drawBar("chartCEE", CEE, "#7c4dff", true);
}

// Mode switching
d3.selectAll(".mode-btn").on("click",function(){
  currentMode = this.dataset.mode;
  d3.selectAll(".mode-btn").classed("active", function(){return this.dataset.mode===currentMode});
  renderAll();
});

renderAll();

// Tooltip delegation
document.addEventListener("mouseover", function(e) {
  const target = d3.select(e.target);
  if(target.node() && target.node().tagName==="rect" && target.attr("fill")!=="#f5f5f5") {
    const d = target.datum();
    if(d && d.name) {
      target.attr("opacity",1);
      d3.select("#tooltip").style("opacity",1).html(`<b>${d.name}</b><br>125: ${d.p125} 篇 (排第${d.rank_125})<br>135: ${d.p135} 篇 (排第${d.rank_135})<br>合计: ${d.p135+d.p125} 篇`)
        .style("left",(e.pageX+14)+"px").style("top",(e.pageY-10)+"px");
    }
  }
});
document.addEventListener("mouseout", function(e) {
  const target = d3.select(e.target);
  if(target.node() && target.node().tagName==="rect" && target.attr("fill")!=="#f5f5f5") {
    target.attr("opacity",0.8);
    d3.select("#tooltip").style("opacity",0);
  }
});
</script>
</body>
</html>"""

html = html.replace("__CN_DATA__", json.dumps(cn_top, ensure_ascii=False))
html = html.replace("__CEE_DATA__", json.dumps(cee_top, ensure_ascii=False))
write_html(os.path.join(OUTPUT_DIR, "institution_visualization.html"), html, "机构排名")
print(f"  中国机构: {len(cn_top)}, 中东欧机构: {len(cee_top)}")
print("Done.")
