"""Offline Frontend Slides player, adapting the upstream fixed-stage architecture.

Original CSS and design references: zarazhangrui/frontend-slides, MIT (2025 Zara Zhang).
Browser edits are plain text, exported explicitly, and never execute document code.
"""
import html
import json
import re

from exports import scene_html
from frontend_slides import ASSETS,edit_bindings,edit_signature,licenses


def json_script(value):
    return json.dumps(value,ensure_ascii=False).replace('<','\\u003c').replace('\u2028','\\u2028').replace('\u2029','\\u2029')


def frontend_html(scenes,theme,fonts,*,preview=False,embedded=False):
    config={'deck_id':edit_signature(scenes),'bindings':edit_bindings(scenes),
            'theme':theme['id'],'count':len(scenes),'preview':preview,'embedded':embedded,
            'notes':[s.get('notes','') for s in scenes]}
    options=''.join(f'<option value="{i}">{i+1:02d} · {html.escape(s["title"])}</option>' for i,s in enumerate(scenes))
    slides=''.join(scene_html(s,interactive=not preview) for s in scenes)
    template=r'''<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="generator" content="SlideCraft Frontend Slides adapter">
<title>__TITLE__</title><style>
/* === OFFLINE GOOGLE FONTS / OPEN FONT LICENSE === */
__FONTS__
/* === UPSTREAM FIXED-STAGE BASE: INCLUDED IN FULL === */
__BASE_CSS__
/* === PRESENTATION THEME === */
:root{--stage-bg:__OUTER__;--slide-bg:__BACKGROUND__;--ease-out-expo:cubic-bezier(.16,1,.3,1)}
*{box-sizing:border-box}body{font-family:'Manrope','SlideCraft CJK',sans-serif}
.deck-viewport{touch-action:pan-y}.deck-stage{outline:none}
.slide{transition:opacity .2s ease,visibility .2s;isolation:isolate}
.reveal{opacity:0;transform:translateY(18px);transition:opacity .24s var(--ease-out-expo),transform .28s var(--ease-out-expo);transition-delay:var(--delay,0ms)}
.slide.visible .reveal{opacity:1;transform:translateY(0)}
.slide.visible{will-change:opacity}.ready .slide.visible .reveal{will-change:opacity,transform}
.instant .slide,.instant .reveal{transition:none!important}
/* === NAVIGATION OUTSIDE THE STAGE === */
.deck-controls{display:flex;align-items:center;gap:8px;max-width:calc(100vw - 24px);padding:8px;border:1px solid #ccd4e5;border-radius:12px;background:#fffffff5;box-shadow:0 4px 24px #0002;color:#17243d}
button,select{font:inherit;font-size:14px;color:#17243d;background:white;border:1px solid #dce2ef;border-radius:6px;height:36px}
button{padding:0 12px;cursor:pointer;white-space:nowrap}button:hover{background:#edf2fc}button:disabled{opacity:.35;cursor:default}
button:focus-visible,select:focus-visible{outline:2px solid #4361ee;outline-offset:2px}select{min-width:70px;max-width:270px;padding:0 8px}
#count{font-size:13px;min-width:45px;font-variant-numeric:tabular-nums;text-align:center}.progress{position:fixed;bottom:0;left:0;height:3px;background:__ACCENT__;width:0;transition:width .2s;z-index:1001}
.status{position:fixed;left:18px;top:12px;right:18px;margin:0;font-size:12px;color:#fff;background:#17243de8;border-radius:5px;padding:7px 10px;z-index:1002;opacity:0;pointer-events:none}.status.show{opacity:1}
.notes{position:fixed;right:18px;bottom:90px;max-width:460px;width:calc(100vw - 36px);max-height:50vh;overflow:auto;white-space:pre-wrap;background:#fffffff8;color:#17243d;border:1px solid #dce2ef;border-radius:8px;padding:18px;z-index:1002;font-size:14px;line-height:1.6;visibility:hidden;opacity:0}.notes.open{visibility:visible;opacity:1}
/* === PLAIN-TEXT EDITING / GRACE-PERIOD HOTZONE === */
.edit-hotzone{position:fixed;top:0;left:0;width:86px;height:72px;z-index:1003}
.edit-toggle{position:fixed;top:12px;left:12px;opacity:0;pointer-events:none;z-index:1004;transition:opacity .16s}
.edit-toggle.show,.edit-toggle.active{opacity:1;pointer-events:auto}
.editing .editable{outline:1px dashed #568bf8;outline-offset:4px;cursor:text}
.editing .editable,.editable.edited{white-space:pre-wrap!important;overflow-wrap:anywhere}
.editable.overfull{outline:3px solid #d22!important;outline-offset:4px}
.edit-actions{position:fixed;top:12px;right:12px;display:flex;gap:6px;z-index:1004;visibility:hidden}
.editing .edit-actions{visibility:visible}
@media(max-width:650px){.deck-controls{gap:5px;padding:6px;bottom:12px}button{font-size:12px;padding:0 8px}select{max-width:130px;font-size:12px}.optional-label{display:none}.edit-actions{top:58px;left:12px;right:auto}}
@media(prefers-reduced-motion:reduce){.slide,.reveal,button,.progress{transition:none!important;animation:none!important}.reveal{transform:none!important}}
@media print{.edit-hotzone,.edit-toggle,.edit-actions,.status,.notes,.progress{display:none!important}.reveal{opacity:1!important;transform:none!important}.editable{outline:none!important}}
body.preview .deck-controls,body.preview .progress,body.preview .edit-hotzone,body.preview .edit-toggle{display:none}
</style></head><body class="instant __BODY_CLASS__">
<!-- Fixed 1920 x 1080 canvas. Fonts, images, styles and controller are embedded. -->
<div class="deck-viewport" id="viewport"><main class="deck-stage" id="deckStage" tabindex="0" aria-label="幻灯片演示">__SLIDES__</main></div>
<nav class="deck-controls" aria-label="演示控制">
<button id="previous" type="button">上一页</button><select id="pages" aria-label="选择演示页面">__OPTIONS__</select><button id="next" type="button">下一页</button>
<span id="count" aria-live="polite"></span><button id="fullscreen" type="button">全屏</button><button id="notes" class="optional-label" type="button">讲稿</button>
</nav>
<div id="progress" class="progress"></div><p id="status" class="status" role="status"></p><aside id="notesPanel" class="notes" aria-label="讲稿备注"></aside>
<div class="edit-hotzone" id="editHotzone"></div><button id="editToggle" class="edit-toggle" type="button" title="编辑文字（E）">编辑文字</button>
<div class="edit-actions"><button id="saveHtml" type="button">保存 HTML</button><button id="saveEdits" type="button">下载修改 JSON</button><button id="clearEdits" type="button">清除本地修改</button></div>
<script type="application/json" id="deck-config">__CONFIG__</script>
<script type="application/json" id="saved-edits">{}</script>
<script type="application/json" id="third-party-licenses">__LICENSES__</script>
<script>
/* === FIXED-STAGE CONTROLLER / NO NETWORK DEPENDENCIES === */
class SlidePresentation{
 constructor(){
  this.config=JSON.parse(document.getElementById('deck-config').textContent);this.stage=document.getElementById('deckStage');this.slides=Array.from(document.querySelectorAll('.slide'));
  this.index=0;this.editing=false;this.edits={};this.notesOpen=false;this.key='slidecraft-frontend:'+this.config.deck_id;
  this.loadEdits();this.setupEditing();this.setupNavigation();this.scale();
  let page=0;try{page=Number(sessionStorage.getItem(this.key+':page'))||0}catch(e){}
  this.show(page);window.addEventListener('resize',()=>this.scale());document.addEventListener('fullscreenchange',()=>this.scale());
  document.fonts.ready.then(()=>{this.stage.querySelectorAll('.slide.visible .edited').forEach(e=>this.checkOverflow(e))});
  requestAnimationFrame(()=>{document.body.classList.remove('instant');document.body.classList.add('ready')});
 }
 scale(){
  const reserve=this.config.preview?0:80;
  const w=window.innerWidth,h=Math.max(1,window.innerHeight-reserve);
  const factor=Math.min(w/1920,h/1080),x=(w-1920*factor)/2,y=(h-1080*factor)/2;
  this.stage.style.transform=`translate(${x}px,${y}px) scale(${factor})`;
 }
 show(value){
  const index=Math.max(0,Math.min(this.slides.length-1,Number(value)||0));this.index=index;
  this.slides.forEach((slide,i)=>{const on=i===index;slide.classList.toggle('active',on);slide.classList.toggle('visible',on);slide.setAttribute('aria-hidden',String(!on));slide.inert=!on});
  document.getElementById('pages').value=String(index);document.getElementById('count').textContent=(index+1)+' / '+this.slides.length;
  document.getElementById('previous').disabled=index===0;document.getElementById('next').disabled=index===this.slides.length-1;
  document.getElementById('progress').style.width=((index+1)/this.slides.length*100)+'%';
  document.getElementById('notesPanel').textContent=this.config.notes[index]||'本页无讲稿。';
  try{sessionStorage.setItem(this.key+':page',String(index))}catch(e){}
 }
 setupNavigation(){
  document.getElementById('previous').addEventListener('click',()=>this.show(this.index-1));document.getElementById('next').addEventListener('click',()=>this.show(this.index+1));
  document.getElementById('pages').addEventListener('change',e=>this.show(e.target.value));
  document.getElementById('notes').addEventListener('click',()=>{this.notesOpen=!this.notesOpen;document.getElementById('notesPanel').classList.toggle('open',this.notesOpen)});
  document.getElementById('fullscreen').addEventListener('click',async()=>{try{if(document.fullscreenElement)await document.exitFullscreen();else await document.documentElement.requestFullscreen()}catch(e){this.notice('当前浏览器限制全屏，可下载 HTML 后在浏览器打开。')}});
  document.addEventListener('keydown',e=>{
   if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'){e.preventDefault();this.saveHtml();return}
   if(e.target.isContentEditable||['INPUT','TEXTAREA','SELECT'].includes(e.target.tagName)||e.ctrlKey||e.metaKey||e.altKey)return;
   if(e.key.toLowerCase()==='e'&&!this.config.preview){e.preventDefault();this.toggleEdit();return}
   const page={ArrowRight:this.index+1,ArrowDown:this.index+1,' ':this.index+1,PageDown:this.index+1,ArrowLeft:this.index-1,ArrowUp:this.index-1,PageUp:this.index-1,Home:0,End:this.slides.length-1}[e.key];
   if(page!==undefined){e.preventDefault();this.show(page)}
  });
  let touch=null;
  this.stage.addEventListener('touchstart',e=>{touch=e.touches.length===1?{x:e.touches[0].clientX,y:e.touches[0].clientY}:null},{passive:true});
  this.stage.addEventListener('touchend',e=>{if(this.editing||!touch||!e.changedTouches.length)return;const t=e.changedTouches[0],dx=t.clientX-touch.x,dy=t.clientY-touch.y;if(Math.abs(dx)>48&&Math.abs(dx)>Math.abs(dy)*1.4)this.show(this.index+(dx<0?1:-1));touch=null},{passive:true});
  let wheel=0;
  document.getElementById('viewport').addEventListener('wheel',e=>{if(this.editing||(this.config.embedded&&!document.fullscreenElement)||Math.abs(e.deltaY)<20||Date.now()-wheel<450)return;e.preventDefault();wheel=Date.now();this.show(this.index+(e.deltaY>0?1:-1))},{passive:false});
 }
 notice(message){const node=document.getElementById('status');node.textContent=message;node.classList.add('show');clearTimeout(this.noticeTimer);this.noticeTimer=setTimeout(()=>node.classList.remove('show'),6000)}
 loadEdits(){
  let saved={};try{saved=JSON.parse(document.getElementById('saved-edits').textContent)}catch(e){}
  try{const local=JSON.parse(localStorage.getItem(this.key)||'{}');saved={...saved,...local}}catch(e){}
  for(const[id,value]of Object.entries(saved)){if(typeof value==='string'&&value.length<=3000&&this.config.bindings[id]){this.edits[id]=value;const el=document.getElementById(id);if(el){el.textContent=value;el.classList.add('edited')}}}
 }
 persist(){try{localStorage.setItem(this.key,JSON.stringify(this.edits))}catch(e){}}
 checkOverflow(el){const over=el.scrollWidth>el.clientWidth+3||el.scrollHeight>el.clientHeight+3;el.classList.toggle('overfull',over);if(over)this.notice('文字超出原区域。请精简，或下载修改 JSON 导回网站重新排版。')}
 toggleEdit(){
  this.editing=!this.editing;document.body.classList.toggle('editing',this.editing);document.getElementById('editToggle').classList.toggle('active',this.editing);
  for(const el of document.querySelectorAll('.editable')){const id=el.dataset.editId,b=this.config.bindings[id];if(!b)continue;el.contentEditable=this.editing?'plaintext-only':'false';
   if(this.editing)el.textContent=this.edits[id]??b.text;
   else{const value=el.textContent;if(value===b.text){delete this.edits[id];el.textContent=b.display;el.classList.remove('edited','overfull')}else{this.edits[id]=value;el.classList.add('edited');this.checkOverflow(el)}}
  }
  this.persist();if(this.editing)this.notice('点击标题或正文编辑。修改在本机暂存；下载修改 JSON 可导回网站并更新 PPTX。')
 }
 setupEditing(){
  const button=document.getElementById('editToggle'),zone=document.getElementById('editHotzone');let timer=null;
  const reveal=()=>{clearTimeout(timer);button.classList.add('show')};const hide=()=>{clearTimeout(timer);timer=setTimeout(()=>{if(!this.editing)button.classList.remove('show')},400)};
  zone.addEventListener('mouseenter',reveal);zone.addEventListener('mouseleave',hide);button.addEventListener('mouseenter',reveal);button.addEventListener('mouseleave',hide);
  zone.addEventListener('click',()=>this.toggleEdit());button.addEventListener('click',()=>this.toggleEdit());
  this.stage.addEventListener('input',e=>{const el=e.target.closest('[data-edit-id]');if(!el||!this.editing)return;const id=el.dataset.editId,value=el.textContent;if(value.length>3000){this.notice('单个段落最多编辑 3000 字，请精简。');return}this.edits[id]=value;this.persist();this.checkOverflow(el)});
  this.stage.addEventListener('paste',e=>{if(!this.editing||!e.target.isContentEditable)return;e.preventDefault();const value=e.clipboardData.getData('text/plain');document.execCommand('insertText',false,value)});
  document.getElementById('saveHtml').addEventListener('click',()=>this.saveHtml());document.getElementById('saveEdits').addEventListener('click',()=>this.saveEdits());
  document.getElementById('clearEdits').addEventListener('click',()=>{this.editing=false;document.body.classList.remove('editing');button.classList.remove('active');this.edits={};try{localStorage.removeItem(this.key)}catch(e){}for(const[id,b]of Object.entries(this.config.bindings)){const el=document.getElementById(id);if(el){el.textContent=b.display;el.contentEditable='false';el.classList.remove('edited','overfull')}}this.notice('已恢复此演示的原始文字。')});
 }
 download(value,name,type){const blob=new Blob([value],{type}),url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=name;link.click();setTimeout(()=>URL.revokeObjectURL(url),1000)}
 saveEdits(){if(this.editing)this.toggleEdit();this.download(JSON.stringify({schema:'slidecraft-frontend-edits/v1',deck_id:this.config.deck_id,edits:this.edits},null,2),'frontend-edits.json','application/json')}
 saveHtml(){
  if(this.editing)this.toggleEdit();const clone=document.documentElement.cloneNode(true);
  clone.querySelector('#saved-edits').textContent=JSON.stringify(this.edits).replace(/</g,'\\u003c');
  clone.querySelector('body').classList.remove('editing');clone.querySelector('#status').classList.remove('show');
  this.download('<!doctype html>\n'+clone.outerHTML,'presentation-edited.html','text/html');
 }
}
window.presentation=new SlidePresentation();
document.querySelectorAll('.reveal').forEach((e,i)=>e.style.setProperty('--delay',(i%4)*45+'ms'));
</script></body></html>'''
    replacements={
        '__TITLE__':html.escape(scenes[0]['title']), '__FONTS__':fonts,
        '__BASE_CSS__':(ASSETS/'viewport-base.css').read_text('utf-8'),
        '__OUTER__':theme['primary'] if theme['id']=='signal' else '#E8EDF2',
        '__BACKGROUND__':theme['background'],'__ACCENT__':theme['accent'],
        '__BODY_CLASS__':'preview' if preview else '', '__SLIDES__':slides,
        '__OPTIONS__':options,'__CONFIG__':json_script(config),'__LICENSES__':json_script(licenses()),
    }
    return re.sub('|'.join(re.escape(k) for k in replacements),lambda m:replacements[m.group()],template).encode('utf-8')
