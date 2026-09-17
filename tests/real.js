'use strict';
global.window=global;global.addEventListener=()=>{};
try{global.navigator.clipboard={};}catch(e){Object.defineProperty(global,'navigator',{value:{clipboard:{}},configurable:true});}
global.localStorage={getItem:()=>null,setItem:()=>{},removeItem:()=>{}};
global.confirm=()=>true;global.prompt=()=>null;global.alert=()=>{};
class El{constructor(tag){this.tagName=tag||'div';this.id='';this._class=new Set();this.style={};this.dataset={};this.children=[];this.parentNode=null;this._listeners={};this.value='';this.textContent='';this.innerText='';this.scrollTop=0;this.scrollLeft=0;this.clientWidth=800;this.clientHeight=600;this.attributes={};this.click=()=>{};this.classList={add:c=>this._class.add(c),remove:c=>this._class.delete(c),toggle:(c,f)=>{if(f===undefined){this._class.has(c)?this._class.delete(c):this._class.add(c);}else f?this._class.add(c):this._class.delete(c);},contains:c=>this._class.has(c)};this.focus=()=>{global.document.activeElement=this;};this.blur=()=>{if(global.document.activeElement===this)global.document.activeElement=null;};this.select=()=>{};this.setAttribute=(k,v)=>{this.attributes[k]=v;if(k==='id')this.id=v;};this.getAttribute=k=>this.attributes[k]??null;this.getBoundingClientRect=()=>({left:0,top:0,width:800,height:600,right:800,bottom:600});this.contains=n=>n===this;this.remove=()=>{if(this.parentNode){const i=this.parentNode.children.indexOf(this);if(i>=0)this.parentNode.children.splice(i,1);this.parentNode=null;}};}
addEventListener(){};set innerHTML(v){this._inner=v;}get innerHTML(){return this._inner||'';}
appendChild(c){this.children.push(c);c.parentNode=this;return c;}
removeChild(c){const i=this.children.indexOf(c);if(i>=0)this.children.splice(i,1);c.parentNode=null;return c;}
insertBefore(c,ref){const i=this.children.indexOf(ref);if(i<0)this.children.unshift(c);else this.children.splice(i,0,c);c.parentNode=this;return c;}
querySelector(){return null;}querySelectorAll(){return [];}}
global.document={activeElement:null,createElement:t=>new El(t),createDocumentFragment:()=>new El('fragment'),getElementById:id=>{const e=new El('div');e.id=id;return e;},addEventListener:()=>{},body:new El('body')};
global.Blob=class{constructor(p){this.p=p;}};
global.URL={createObjectURL:()=>'blob:fake',revokeObjectURL:()=>{}};
global.FileReader=class{};
const fs=require('fs'),vm=require('vm');
global.XLSX=require(process.env.SHEETJS);
global.XLSXF=require(process.env.OOXMLJS);
vm.runInThisContext(fs.readFileSync(process.env.APPJS,'utf8'),{filename:process.env.APPJS});
const files=process.argv.slice(2);
let failed=0;
for(const f of files){
  try{
    const buf=fs.readFileSync(f);
    const data=new Uint8Array(buf.buffer,buf.byteOffset,buf.byteLength);
    const wb=XLSX.read(data,{type:'array',cellDates:true,cellNF:true});
    refreshStateFromXLSX(wb,f,data);
    let cells=0;for(const k in sheets){cells+=Object.keys(sheets[k].cells).length;}
    console.log('OK', f.split('/').pop(), '| sheets:', order.join(','), '| cells:', cells, '| OO:', OO!==null, '| ooSheets:', Object.keys(sheets).filter(n=>sheets[n].oo).join(',')||'(none)');
  }catch(e){
    failed++;
    console.log('FAIL', f.split('/').pop(), '->', e.message);
  }
}
console.log(failed?('TOTAL FAILED: '+failed):'ALL LOADED');
