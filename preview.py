"""Client-side slide navigation: no Streamlit rerun or network request on page turns."""
import html
import json
import re

from exports import scene_html
from layouts import LAYOUT_LABELS


def preview_html(scenes, deck_key):
    if not scenes:
        raise ValueError("预览至少需要一页幻灯片。")
    panels = []
    options = []
    labels = []
    for i, scene in enumerate(scenes):
        panels.append(f'<article class="panel" aria-hidden="true" inert data-page="{i}">{scene_html(scene)}</article>')
        options.append(f'<option value="{i}">{i+1:02d} · {html.escape(scene["title"])}</option>')
        label = LAYOUT_LABELS.get(scene["layout"], "封面")
        reason = scene.get("layout_reason", "")
        labels.append(label + (" · " + reason if reason else ""))
    data = json.dumps({"key": "slidecraft-preview:" + str(deck_key), "labels": labels,
                       "count": len(scenes), "height": scenes[0]["height"]}, ensure_ascii=False)
    # Escape script delimiters even when a document title contains arbitrary markup.
    data = data.replace("<", "\\u003c").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    template = r'''<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{box-sizing:border-box}html,body{margin:0;width:100%;height:100%;overflow:hidden}
body{font-family:"Microsoft YaHei",Arial,sans-serif;color:#17243D}
.viewer{height:100%;display:flex;flex-direction:column;gap:12px;outline:none}
.toolbar{display:flex;align-items:center;gap:10px;flex-shrink:0;padding:2px}
button,select{font:inherit;font-size:15px;height:40px;border:1px solid #DCE2EF;background:white;color:#17243D;border-radius:8px}
button{padding:0 15px;cursor:pointer;white-space:nowrap;transition:background .15s,border-color .15s}
button:hover:enabled{background:#EAF0FC;border-color:#2147B1}button:disabled{opacity:.4;cursor:default}
button:focus-visible,select:focus-visible,.viewer:focus-visible{outline:2px solid #2147B1;outline-offset:2px}
select{flex:1;min-width:0;padding:0 10px;text-overflow:ellipsis}.counter{font-size:14px;font-variant-numeric:tabular-nums;white-space:nowrap;color:#62708A;min-width:48px;text-align:right}
.stage-wrap{flex:1;min-height:0;display:flex;align-items:flex-start;justify-content:center;padding:2px}
.stage{position:relative;overflow:hidden;isolation:isolate;touch-action:pan-y;background:white;border-radius:10px;box-shadow:0 2px 12px #17243d0b}
.panel{position:absolute;inset:0;opacity:0;visibility:hidden;pointer-events:none;transform:translate3d(var(--enter,12px),0,0);transition:opacity 170ms ease,transform 190ms cubic-bezier(.2,.65,.3,1),visibility 190ms}
.panel.active,.panel.leaving{will-change:opacity,transform}
.panel.active{opacity:1;visibility:visible;pointer-events:auto;transform:translate3d(0,0,0);z-index:2}
.panel.leaving{visibility:visible;transform:translate3d(var(--exit,-12px),0,0);z-index:1}
.slide{position:absolute;left:0;top:0;transform-origin:top left;overflow:hidden}
.caption{flex-shrink:0;margin:0 2px;min-height:38px;font-size:13px;line-height:1.5;color:#62708A}
.hint{margin:0 2px;font-size:12px;color:#62708A;flex-shrink:0}
.initializing .stage{visibility:hidden}.instant .panel{transition:none}
@media(max-width:520px){.toolbar{gap:6px}button{padding:0 9px;font-size:14px}.counter{display:none}select{font-size:14px}.viewer{gap:8px}.caption{font-size:12px}}
@media(prefers-reduced-motion:reduce){.panel,button{transition:none}.panel{transform:none}}
</style></head><body class="initializing">
<main class="viewer" tabindex="0" aria-label="幻灯片预览，可用左右方向键翻页">
<nav class="toolbar" aria-label="预览翻页">
<button id="previous" type="button" aria-label="上一页">上一页</button>
<select id="page-select" aria-label="选择预览页面">__OPTIONS__</select>
<button id="next" type="button" aria-label="下一页">下一页</button>
<span id="counter" class="counter" aria-live="polite" aria-atomic="true"></span>
</nav>
<div class="stage-wrap"><div id="stage" class="stage">__PANELS__</div></div>
<p id="caption" class="caption"></p>
<p class="hint">点击上一页 / 下一页，或在预览中用左右方向键；手机可左右滑动。</p>
</main>
<script>
(()=>{
const config=__DATA__;
const viewer=document.querySelector('.viewer'),stage=document.getElementById('stage');
const panels=Array.from(document.querySelectorAll('.panel'));
const select=document.getElementById('page-select'),previous=document.getElementById('previous'),next=document.getElementById('next');
let current=-1;
function fit(){
 const box=document.querySelector('.stage-wrap');
 const width=Math.max(1,Math.min(box.clientWidth-4,(box.clientHeight-4)*1280/config.height));
 stage.style.width=width+'px';stage.style.height=(width*config.height/1280)+'px';
 panels.forEach(p=>{p.firstElementChild.style.transform='scale('+(width/1280)+')'});
}
function readPosition(){
 try{const v=Number(sessionStorage.getItem(config.key));return Number.isInteger(v)?v:0}catch(e){return 0}
}
function show(value,animate=true){
 const index=Math.max(0,Math.min(config.count-1,Number(value)||0));
 if(index===current)return;
 const direction=current<0||index>current?1:-1;
 viewer.classList.toggle('instant',!animate);
 stage.style.setProperty('--enter',(direction*12)+'px');
 stage.style.setProperty('--exit',(-direction*12)+'px');
 const old=current;
 panels.forEach((p,i)=>{
   p.classList.toggle('leaving',i===old&&animate);
   p.classList.toggle('active',i===index);
   p.setAttribute('aria-hidden',i===index?'false':'true');
   p.inert=i!==index;
 });
 current=index;select.value=String(index);
 previous.disabled=index===0;next.disabled=index===config.count-1;
 document.getElementById('counter').textContent=(index+1)+' / '+config.count;
 document.getElementById('caption').textContent='本页版式：'+config.labels[index];
 try{sessionStorage.setItem(config.key,String(index))}catch(e){}
}
previous.addEventListener('click',()=>show(current-1));
next.addEventListener('click',()=>show(current+1));
select.addEventListener('change',()=>show(select.value));
viewer.addEventListener('keydown',e=>{
 if(e.altKey||e.ctrlKey||e.metaKey||['SELECT','INPUT','TEXTAREA'].includes(e.target.tagName))return;
 const target={ArrowLeft:current-1,ArrowRight:current+1,Home:0,End:config.count-1}[e.key];
 if(target!==undefined){e.preventDefault();show(target)}
});
let touch=null;
stage.addEventListener('touchstart',e=>{touch=e.touches.length===1?{x:e.touches[0].clientX,y:e.touches[0].clientY}:null},{passive:true});
stage.addEventListener('touchend',e=>{
 if(!touch||!e.changedTouches.length)return;
 const dx=e.changedTouches[0].clientX-touch.x,dy=e.changedTouches[0].clientY-touch.y;
 if(Math.abs(dx)>48&&Math.abs(dx)>Math.abs(dy)*1.4)show(current+(dx<0?1:-1));
 touch=null;
},{passive:true});
stage.addEventListener('touchcancel',()=>{touch=null},{passive:true});
new ResizeObserver(fit).observe(document.querySelector('.stage-wrap'));
window.addEventListener('resize',fit);
show(readPosition(),false);fit();document.body.classList.remove('initializing');
requestAnimationFrame(()=>viewer.classList.remove('instant'));
})();
</script></body></html>'''
    replacements={"__OPTIONS__":"".join(options),"__PANELS__":"".join(panels),"__DATA__":data}
    return re.sub(r"__OPTIONS__|__PANELS__|__DATA__",lambda m:replacements[m.group()],template)
